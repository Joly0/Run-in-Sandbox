<#
.SYNOPSIS
    Update management module for Run-in-Sandbox

.DESCRIPTION
    This module provides update management functionality for the Run-in-Sandbox application.
    It handles checking for updates, downloading updates, and applying them.
#>

# Check for updates and return current/latest versions with branch info
function Get-VersionInfo {
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
    
    $Current = $null
    $CurrentBranch = "master"
    $Latest = $null
    $LatestBranch = "master"
    
    # Get current version from version.json
    $VersionJsonFile = "$Run_in_Sandbox_Folder\version.json"
    if (Test-Path $VersionJsonFile) {
        try {
            $VersionData = Get-Content $VersionJsonFile -Raw | ConvertFrom-Json
            $Current = $VersionData.version
            $CurrentBranch = if ($VersionData.branch) { $VersionData.branch } else { "master" }
        } catch {}
    }
    
    # Fallback to config XML if version.json doesn't exist
    if (-not $Current) {
        try {
            $Config = [xml](Get-Content $XML_Config)
            $Current = $Config.Configuration.CurrentVersion
        } catch {}
    }
    
    # Get latest version from GitHub (use current branch for update channel)
    try {
        $LatestUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$CurrentBranch/version.json"
        $LatestData = Invoke-RestMethod -Uri $LatestUrl -UseBasicParsing -TimeoutSec 5
        $Latest = $LatestData.version
        $LatestBranch = if ($LatestData.branch) { $LatestData.branch } else { $CurrentBranch }
        Write-LogMessage -Message_Type "INFO" -Message "[UPDATE] Latest version on '$CurrentBranch': $Latest"
    } catch { Write-LogMessage -Message_Type "ERROR" -Message "[UPDATE] Could not determine latest Version from github: $($_.Exception.Message)" }
    
    return @{
        Current = $Current
        CurrentBranch = $CurrentBranch
        Latest = $Latest
        LatestBranch = $LatestBranch
    }
}

# Main update check function (called on sandbox launch)
function Start-UpdateCheck {
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
    
    Write-LogMessage -Message_Type "INFO" -Message "[UPDATE] Checking for updates..."
    
    # Skip if dismissed today
    $DismissFile = "$Run_in_Sandbox_Folder\temp\DismissedUntil.txt"
    if (Test-Path $DismissFile) {
        try {
            if ((Get-Content $DismissFile -Raw).Trim() -eq (Get-Date).ToString('yyyy-MM-dd')) { return }
        } catch { Remove-Item $DismissFile -Force -ErrorAction SilentlyContinue }
    }
    
    # Check network connectivity
    try { $null = Test-Connection -ComputerName "raw.githubusercontent.com" -Count 1 -Quiet } catch { return }
    
    # Get versions and compare
    $VersionInfo = Get-VersionInfo
    if (-not $VersionInfo.Current -or -not $VersionInfo.Latest) { Write-LogMessage -Message_Type "ERROR" -Message "[UPDATE] Could not determine current or latest Version"; return }
    
    try {
        $CurrentDate = [DateTime]::ParseExact($VersionInfo.Current, 'yyyy-MM-dd', $null)
        $LatestDate = [DateTime]::ParseExact($VersionInfo.Latest, 'yyyy-MM-dd', $null)
        
        if ($LatestDate -gt $CurrentDate) {
            Write-LogMessage -Message_Type "INFO" -Message "[UPDATE] Update available: $($VersionInfo.Current) -> $($VersionInfo.Latest)"
            $Config = [xml](Get-Content $XML_Config)
            Show-UpdateToast -LatestVersion $VersionInfo.Latest -Language $Config.Configuration.Main_Language
        } else {
            Write-LogMessage -Message_Type "INFO" -Message "[UPDATE] No new Version found"
        }
    } catch {}
}

# Get changelog for specific version from GitHub
function Get-ChangelogForVersion {
    param([string]$Version, [string]$Branch = "master")
    try {
        $Content = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$Branch/CHANGELOG.md" -UseBasicParsing -TimeoutSec 10
        $Lines = $Content -split "`n"
        $StartIndex = -1
        for ($i = 0; $i -lt $Lines.Count; $i++) {
            if ($Lines[$i] -match "^##\s+$Version") { $StartIndex = $i; break }
        }
        
        if ($StartIndex -eq -1) {
            # Version not found, return full changelog with a note
            $Note = "No specific changes found for version $Version. Showing full changelog:`n`n"
            return $Note + $Content
        }
        
        $ChangelogSection = @()
        for ($i = $StartIndex; $i -lt $Lines.Count; $i++) {
            if ($i -ne $StartIndex -and $Lines[$i] -match '^##\s+\d{4}-\d{2}-\d{2}') { break }
            $ChangelogSection += $Lines[$i]
        }
        
        $Result = ($ChangelogSection -join "`n").Trim()
        if ([string]::IsNullOrWhiteSpace($Result)) {
            # Empty section found, return full changelog with a note
            $Note = "No changes recorded for version $Version. Showing full changelog:`n`n"
            return $Note + $Content
        }
        
        return $Result
    } catch {
        return "Unable to fetch changelog from GitHub. Please visit https://github.com/Joly0/Run-in-Sandbox/blob/$Branch/CHANGELOG.md"
    }
}

# Show Windows toast notification for available update
function Show-UpdateToast {
    param([string]$LatestVersion, [string]$Language = "en-US")
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $Strings = Get-LocalizedUpdateStrings -Language $Language
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
    
    $IconPath = "$Run_in_Sandbox_Folder\sandbox.ico"
    # Create a simple batch file to launch the dialog
    $HandlerScript = "$Run_in_Sandbox_Folder\temp\UpdateClickHandler.ps1"
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    if (-not (Test-Path $TempFolder)) { New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null }
    
    # Clean up any existing scheduled task from previous versions
    try {
        $Task = Get-ScheduledTask -TaskName "RunInSandboxUpdateHandler" -ErrorAction SilentlyContinue
        if ($Task) {
            Unregister-ScheduledTask -TaskName "RunInSandboxUpdateHandler" -Confirm:$false
        }
    } catch {
        # Ignore errors if task doesn't exist or can't be removed
    }
    
    $ScriptContent = @"
`$Run_in_Sandbox_Folder = "$Run_in_Sandbox_Folder"
try {
    . "`$Run_in_Sandbox_Folder\CommonFunctions.ps1"
    Add-Type -AssemblyName PresentationFramework
    
    # Check if we're in STA mode
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        # Not in STA mode, start new PowerShell process
        `$Command = "'`$Run_in_Sandbox_Folder = '$Run_in_Sandbox_Folder'; . '`$Run_in_Sandbox_Folder\CommonFunctions.ps1'; Show-ChangelogDialog -LatestVersion '$LatestVersion'"
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -Command `$Command"
        exit
    }
    
    Show-ChangelogDialog -LatestVersion "$LatestVersion"
} catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("Error: `$(`$_.Exception.Message)", "Error", "OK", "Error")
}
"@
    $ScriptContent | Set-Content -Path $HandlerScript
    
    # Create a batch file to launch the PowerShell script
    $BatchFile = "$TempFolder\ShowUpdateDialog.bat"
    @"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "$HandlerScript"
exit
"@ | Set-Content -Path $BatchFile
    
    # Register protocol handler for toast clicks
    $ProtocolKey = "HKCU:\Software\Classes\run-in-sandbox-update"
    if (-not (Test-Path $ProtocolKey)) {
        New-Item -Path $ProtocolKey -Force | Out-Null
        Set-ItemProperty -Path $ProtocolKey -Name "(Default)" -Value "URL:Run-in-Sandbox Update Protocol"
        Set-ItemProperty -Path $ProtocolKey -Name "URL Protocol" -Value ""
        $CommandKey = "$ProtocolKey\shell\open\command"
        New-Item -Path $CommandKey -Force | Out-Null
    } else {
        $CommandKey = "$ProtocolKey\shell\open\command"
    }
    # Always update the command to use the batch file
    Set-ItemProperty -Path $CommandKey -Name "(Default)" -Value "`"$BatchFile`""
    
    $ToastXml = @"
<toast activationType="protocol" launch="run-in-sandbox-update:$LatestVersion">
    <visual><binding template="ToastGeneric">
        <text>$($Strings.ToastTitle)</text>
        <text>$($Strings.ToastMessage -f $LatestVersion)</text>
        <image placement="appLogoOverride" src="file:///$($IconPath.Replace('\', '/'))"/>
    </binding></visual>
    <audio src="ms-winsoundevent:Notification.Default"/>
</toast>
"@
    
    try {
        # Create a temporary AppUserModelID if needed
        $AppId = "Run-in-Sandbox"
        $RegPath = "HKCU:\Software\Classes\AppUserModelId\$AppId"
        if (-not (Test-Path $RegPath)) {
            New-Item -Path $RegPath -Force | Out-Null
            Set-ItemProperty -Path $RegPath -Name "DisplayName" -Value "Run-in-Sandbox"
        }
        
        $XmlDoc = [Windows.Data.Xml.Dom.XmlDocument]::new()
        $XmlDoc.LoadXml($ToastXml)
        $Toast = [Windows.UI.Notifications.ToastNotification]::new($XmlDoc)
        $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId)
        $Notifier.Show($Toast)
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "[UPDATE] Failed to show toast notification: $($_.Exception.Message)"
    }
}

# Show WPF dialog with changelog and update options
function Show-ChangelogDialog {
    param([string]$LatestVersion)
    
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
    
    $VersionInfo = Get-VersionInfo
    $Changelog = Get-ChangelogForVersion -Version $LatestVersion -Branch $VersionInfo.CurrentBranch
    $Strings = Get-LocalizedUpdateStrings -Language ([xml](Get-Content $XML_Config)).Configuration.Main_Language
    Add-Type -AssemblyName PresentationFramework
    
    # Load XAML from file and replace placeholders
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_UpdateDialog.xaml"
    
    if (-not (Test-Path $XamlPath)) {
        [System.Windows.MessageBox]::Show("XAML file not found: $XamlPath", "Error", "OK", "Error")
        return
    }
    
    $Xaml = Get-Content $XamlPath -Raw
    
    # Save changelog to a temporary file and load it from file
    $TempChangelogFile = "$Run_in_Sandbox_Folder\temp\changelog_temp.txt"
    $Changelog | Set-Content -Path $TempChangelogFile -Encoding UTF8
    
    $Xaml = $Xaml -replace 'PLACEHOLDER_TITLE', ($Strings.DialogTitle -f $LatestVersion)
    $Xaml = $Xaml -replace 'PLACEHOLDER_CURRENT_LABEL', "$($Strings.CurrentVersionLabel): "
    $Xaml = $Xaml -replace 'PLACEHOLDER_CURRENT_VERSION', $VersionInfo.Current
    $Xaml = $Xaml -replace 'PLACEHOLDER_LATEST_LABEL', "$($Strings.LatestVersionLabel): "
    $Xaml = $Xaml -replace 'PLACEHOLDER_LATEST_VERSION', $LatestVersion
    $Xaml = $Xaml -replace 'PLACEHOLDER_CHANGELOG', $TempChangelogFile
    $Xaml = $Xaml -replace 'PLACEHOLDER_FOOTER', $Strings.DialogFooter
    $Xaml = $Xaml -replace 'PLACEHOLDER_DISMISS_BUTTON', $Strings.DismissButton
    $Xaml = $Xaml -replace 'PLACEHOLDER_UPDATE_BUTTON', $Strings.UpdateButton
    
    try {
        $Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$Xaml)
        $Window = [Windows.Markup.XamlReader]::Load($Reader)
        
        $UpdateButton = $Window.FindName("UpdateButton")
        $DismissButton = $Window.FindName("DismissButton")
        $ChangelogText = $Window.FindName("ChangelogText")
        
        # Load changelog from file
        if ($ChangelogText -and (Test-Path $TempChangelogFile)) {
            $ChangelogText.Text = Get-Content -Path $TempChangelogFile -Raw -Encoding UTF8
        }
        
        $UpdateButton.Add_Click({
            # Save decision to update
            $TempFolder = "$Run_in_Sandbox_Folder\temp"
            if (-not (Test-Path $TempFolder)) { New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null }
            @{ action = "update"; latestVersion = $LatestVersion } | ConvertTo-Json | Set-Content "$TempFolder\UpdateState.json"
            [System.Windows.MessageBox]::Show($Strings.UpdateScheduledMessage, $Strings.DialogTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $Window.Close()
        })
        
        $DismissButton.Add_Click({
            # Dismiss for today
            $TempFolder = "$Run_in_Sandbox_Folder\temp"
            if (-not (Test-Path $TempFolder)) { New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null }
            Set-Content -Path "$TempFolder\DismissedUntil.txt" -Value (Get-Date).ToString('yyyy-MM-dd')
            $Window.Close()
        })
        
        $Window.ShowDialog() | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show("Error showing dialog: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}

# Load localized strings for update UI
function Get-LocalizedUpdateStrings {
    param([string]$Language = "en-US")
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $LanguageFile = "$Run_in_Sandbox_Folder\Languages_XML\Language_$Language.xml"
    
    # Fallback to en-US if language file not found
    if (-not (Test-Path $LanguageFile)) {
        $LanguageFile = "$Run_in_Sandbox_Folder\Languages_XML\Language_en-US.xml"
    }
    
    $LangXml = [xml](Get-Content $LanguageFile)
    $UpdateStrings = $LangXml.Configuration.Update
    return @{
        ToastTitle = $UpdateStrings.ToastTitle
        ToastMessage = $UpdateStrings.ToastMessage
        DialogTitle = $UpdateStrings.DialogTitle
        CurrentVersionLabel = $UpdateStrings.CurrentVersionLabel
        LatestVersionLabel = $UpdateStrings.LatestVersionLabel
        DialogFooter = $UpdateStrings.DialogFooter
        UpdateButton = $UpdateStrings.UpdateButton
        DismissButton = $UpdateStrings.DismissButton
        UpdateScheduledMessage = $UpdateStrings.UpdateScheduledMessage
        UpdateSuccessTitle = $UpdateStrings.UpdateSuccessTitle
        UpdateSuccessMessage = $UpdateStrings.UpdateSuccessMessage
        RollbackTitle = $UpdateStrings.RollbackTitle
        RollbackMessage = $UpdateStrings.RollbackMessage
        ErrorTitle = $UpdateStrings.ErrorTitle
        ViewChangesButton = if ($UpdateStrings.ViewChangesButton) { $UpdateStrings.ViewChangesButton } else { "View Changes" }
    }
}

# Get saved update decision from previous user interaction
function Get-UpdateDecision {
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $StateFile = "$Run_in_Sandbox_Folder\temp\UpdateState.json"
    if (Test-Path $StateFile) {
        try { return Get-Content $StateFile -Raw | ConvertFrom-Json } catch { return $null }
    }
    return $null
}

Export-ModuleMember -Function @(
    'Get-VersionInfo',
    'Start-UpdateCheck',
    'Get-ChangelogForVersion',
    'Show-UpdateToast',
    'Show-ChangelogDialog',
    'Get-LocalizedUpdateStrings',
    'Get-UpdateDecision'
)