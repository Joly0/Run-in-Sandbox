param (
    [Parameter(Mandatory=$true)] [String]$Type,
    [Parameter(Mandatory=$true)] [String]$ScriptPath
)

#Start-Transcript -Path $(Join-Path -Path $([System.Environment]::GetEnvironmentVariables('Machine').TEMP) -ChildPath "RunInSandbox.log")

$special_char_array = 'é', 'è', 'à', 'â', 'ê', 'û', 'î', 'ä', 'ë', 'ü', 'ï', 'ö', 'ù', 'ò', '~', '!', '@', '#', '$', '%', '^', '&', '+', '=', '}', '{', '|', '<', '>', ';'
foreach ($char in $special_char_array) {
    if ($ScriptPath -like "*$char*") {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $message = "There is a special character in the path of the file (`'" + $char + "`').`nWindows Sandbox does not support this!"
        [System.Windows.Forms.MessageBox]::Show($message, "Issue with your file", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        EXIT
    }
}

$ScriptPath = $ScriptPath.replace('"', '')
$ScriptPath = $ScriptPath.Trim();
$ScriptPath = [WildcardPattern]::Escape($ScriptPath)

if ( ($Type -eq "Folder_Inside") -or ($Type -eq "Folder_On") ) {
    $DirectoryName = (Get-Item $ScriptPath).fullname
} else {
    $FolderPath = Split-Path (Split-Path "$ScriptPath" -Parent) -Leaf
    $DirectoryName = (Get-Item $ScriptPath).DirectoryName
    $FileName = (Get-Item $ScriptPath).BaseName
    $Full_FileName = (Get-Item $ScriptPath).Name
}

$Sandbox_Desktop_Path = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Shared_Path = "$Sandbox_Desktop_Path\$FolderPath"

$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$Full_Startup_Path = "$Sandbox_Shared_Path\$Full_FileName"
$Full_Startup_Path_Quoted = """$Full_Startup_Path"""

$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

# Import required modules
Import-Module "$Run_in_Sandbox_Folder\Modules\Shared\Logging.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Shared\Config.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\SevenZip.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\WSB.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\StartupScripts.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\Update.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\UI.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\Dialogs.psm1" -Force

# Start asynchronous update check (non-blocking)
Start-Job -ScriptBlock {
    param($Run_in_Sandbox_Folder)
    Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\Update.psm1" -Force
    Start-UpdateCheck
} -ArgumentList $Run_in_Sandbox_Folder | Out-Null

$xml = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
$my_xml = [xml](Get-Content $xml)
$Sandbox_VGpu = $my_xml.Configuration.VGpu
$Sandbox_Networking = $my_xml.Configuration.Networking
$Sandbox_ReadOnlyAccess = $my_xml.Configuration.ReadOnlyAccess
$Sandbox_WSB_Location = $my_xml.Configuration.WSB_Location
$Sandbox_AudioInput = $my_xml.Configuration.AudioInput
$Sandbox_VideoInput = $my_xml.Configuration.VideoInput
$Sandbox_ProtectedClient = $my_xml.Configuration.ProtectedClient
$Sandbox_PrinterRedirection = $my_xml.Configuration.PrinterRedirection
$Sandbox_ClipboardRedirection = $my_xml.Configuration.ClipboardRedirection
$Sandbox_MemoryInMB = $my_xml.Configuration.MemoryInMB
$WSB_Cleanup = $my_xml.Configuration.WSB_Cleanup
$Hide_Powershell = $my_xml.Configuration.Hide_Powershell

[System.Collections.ArrayList]$PowershellParameters = @(
    '-sta'
    '-WindowStyle'
    'Hidden'
    '-NoProfile'
    '-ExecutionPolicy'
    'Unrestricted'
)

if ($Hide_Powershell -eq "False") {
    $PowershellParameters[[array]::IndexOf($PowershellParameters, "Hidden")] = "Normal"
}

$PSRun_File = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $PowershellParameters -File"
$PSRun_Command = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $PowershellParameters -Command"

if ($Sandbox_WSB_Location -eq "Default") {
    $Sandbox_File_Path = "$env:temp\$FileName.wsb"
} else {
    $Sandbox_File_Path = "$Sandbox_WSB_Location\$FileName.wsb"
}

if (Test-Path $Sandbox_File_Path) {
    Remove-Item $Sandbox_File_Path
}


switch ($Type) {
    "7Z" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()
        
        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path $Host7ZipPath -Parent
            
            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }
            
            $Script:Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File"
            
            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Ensure-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }
            
            $CachedInstaller = "$Sandbox_Root_Path\temp\7zSetup.msi"
            
            # Install 7-Zip in sandbox then extract
            $Script:Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File' -Wait`""
            
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "CMD" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "EXE" {
        $DialogResult = Show-EXEDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_Quoted $Full_Startup_Path_Quoted
        $EXE_Command_File = "$Run_in_Sandbox_Folder\EXE_Command_File.txt"
        
        $EXE_Installer = "$Sandbox_Root_Path\EXE_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$EXE_Installer"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Script:Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "Folder_On" {
        $Startup_Command = Enable-StartupScripts -OriginalCommand ""
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "Folder_Inside" {
        $Startup_Command = Enable-StartupScripts -OriginalCommand ""
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "HTML" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "URL" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Sandbox_Root_Path"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "Intunewin" {
        $Intunewin_Folder = "C:\IntuneWin\$FileName.intunewin"
        $Intunewin_Content_File = "$Run_in_Sandbox_Folder\Intunewin_Folder.txt"
        $Intunewin_Command_File = "$Run_in_Sandbox_Folder\Intunewin_Install_Command.txt"
        $Intunewin_Folder | Out-File $Intunewin_Content_File -Force -NoNewline

        $DialogResult = Show-IntunewinDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -FileName $FileName

        $Intunewin_Installer = "$Sandbox_Root_Path\IntuneWin_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$Intunewin_Installer"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Script:Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "ISO" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()
        
        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path $Host7ZipPath -Parent
            
            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }
            
            $Script:Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO"
            
            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation for ISO: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Ensure-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }
            
            $CachedInstaller = "$Run_in_Sandbox_Folder\temp\7zSetup.msi"
            
            # Install 7-Zip in sandbox then extract ISO
            $Script:Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO' -Wait`""
            
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer for ISO: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "MSI" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")
        $DialogResult = Show-MSIDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_UnQuoted $Full_Startup_Path_UnQuoted
        $Script:Startup_Command = $DialogResult

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Script:Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "MSIX" {
        $Script:Startup_Command = $PSRun_Command + " " + "Add-AppPackage -Path $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PDF" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PPKG" {
        $Script:Startup_Command = $PSRun_Command + " " + "Install-ProvisioningPackage $Full_Startup_Path_Quoted -forceinstall -quietinstall"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PS1Basic" {
        $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PS1System" {
        $Script:Startup_Command = "$Sandbox_Root_Path\PsExec.exe \\localhost -nobanner -accepteula -s Powershell -ExecutionPolicy Bypass -File $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PS1Params" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")
        $DialogResult = Show-ParamsDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_UnQuoted $Full_Startup_Path_UnQuoted -PSRun_File $PSRun_File
        $Script:Startup_Command = $DialogResult
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Script:Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "REG" {
        $Script:Startup_Command = "REG IMPORT $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "SDBApp" {
        $AppBundle_Installer = "$Sandbox_Root_Path\AppBundle_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$AppBundle_Installer"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "VBSBasic" {
        $Script:Startup_Command = "wscript.exe $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "VBSParams" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $DialogResult = Show-VBSParamsDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_UnQuoted $Full_Startup_Path_UnQuoted
        $Script:Startup_Command = $DialogResult

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Script:Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "ZIP" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Expand-Archive -LiteralPath '$Full_Startup_Path' -DestinationPath '$Sandbox_Desktop_Path\ZIP_extracted'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
}

Start-Process -FilePath $Sandbox_File_Path -Wait
do {
    Start-Sleep -Seconds 1
} while (Get-Process -Name "WindowsSandboxServer" -ErrorAction SilentlyContinue)

if ($WSB_Cleanup -eq $True) {
    Remove-Leftovers -RemovalPath $Sandbox_File_Path
    Remove-Leftovers -RemovalPath $Intunewin_Command_File
    Remove-Leftovers -RemovalPath $Intunewin_Content_File
    Remove-Leftovers -RemovalPath $EXE_Command_File
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\App_Bundle.sdbapp"
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\NotepadPayload"
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\startup-scripts\OriginalCommand.txt"
}

# Check for pending update installation (runs regardless of cleanup setting)
$UpdateDecision = Get-UpdateDecision
if ($UpdateDecision -and $UpdateDecision.action -eq "update") {
    Write-Host "Installing update to version $($UpdateDecision.latestVersion)..." -ForegroundColor Green
    Write-Host "Starting update process in new window..." -ForegroundColor Cyan
    
    # Download and execute Install_Run-in-Sandbox.ps1 as a separate process
    try {
        $VersionInfo = Get-VersionInfo
        $Branch = $VersionInfo.CurrentBranch
        $InstallScriptUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$Branch/Install_Run-in-Sandbox.ps1"
        
        # Download to temp file
        $TempInstaller = "$env:TEMP\Run-in-Sandbox-Updater.ps1"
        Invoke-WebRequest -Uri $InstallScriptUrl -OutFile $TempInstaller -UseBasicParsing
        
        # Launch installer in new elevated process (don't wait)
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$TempInstaller`" -AutoUpdate -NoCheckpoint" -Verb RunAs
        
        Write-Host "Update process started. This window will now close." -ForegroundColor Green
        Write-Host "The update will complete in the new window." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        # Exit immediately to release file locks
        exit 0
    } catch {
        Write-Host "Failed to start update process: $_" -ForegroundColor Red
        Write-Host "You can manually update by running: powershell -c `"irm https://run-in-sandbox.com/install.ps1 | iex`"" -ForegroundColor Yellow
    }
}