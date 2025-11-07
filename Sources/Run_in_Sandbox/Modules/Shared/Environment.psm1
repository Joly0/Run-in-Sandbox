<#
.SYNOPSIS
    Environment and file system module for Run-in-Sandbox

.DESCRIPTION
    This module provides environment detection and file system management functionality 
    for the Run-in-Sandbox application. It handles admin checks, source validation, 
    and file operations.
#>

# Global variables
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption
$currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
if (Test-Path -LiteralPath "Registry::HKEY_USERS\$currentSid\Volatile Environment" -ErrorAction SilentlyContinue) {
    $Current_User_SID = $currentSid
    $HKCU = "Registry::HKEY_USERS\$Current_User_SID"
    $HKCU_Classes = "Registry::HKEY_USERS\${Current_User_SID}_Classes"
} else {
    $HKCU = 'HKCU:'
    $HKCU_Classes = 'HKCU:\Software\Classes'
}
$Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
$Sources = $Current_Folder + "\" + "Sources\*"

# Load Windows Forms assembly for message boxes
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

function Test-IsAdmin {
    $windowsIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $windowsPrincipal = [Security.Principal.WindowsPrincipal]::new($windowsIdentity)
    return $windowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Admin {
    param(
        [string]$EffectiveBranch,
        [switch]$NoCheckpoint,
        [switch]$DeepClean,
        [switch]$AutoUpdate
    )
    if (Test-IsAdmin) {
        Write-Verbose "Already elevated."
        return
    }

    Write-Info "Elevation required. Restarting installer as Administrator..." ([ConsoleColor]::Yellow)

    # Simplified, reliable elevation: download current branch installer to temp and re-run with same parameters.
    $TempScript = Join-Path ([IO.Path]::GetTempPath()) "Install_Run-in-Sandbox.Elevated.ps1"
    $InstallerUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$EffectiveBranch/Install_Run-in-Sandbox.ps1"

    try {
        Write-Verbose "Downloading elevated installer from: $InstallerUrl"
        Invoke-WebRequest -Uri $InstallerUrl -UseBasicParsing -OutFile $TempScript -TimeoutSec 45 -ErrorAction Stop
    } catch {
        Write-Info "Failed to download installer for elevation: $($_.Exception.Message)" ([ConsoleColor]::Red)
        Read-Host "Press Enter to exit"
        break script
    }

    $argsList = @(
        "-NoExit",
        "-ExecutionPolicy", "Bypass",
        "-File", $TempScript
    )
    if ($NoCheckpoint) { $argsList += "-NoCheckpoint" }
    if ($DeepClean)    { $argsList += "-DeepClean" }
    if ($AutoUpdate)   { $argsList += "-AutoUpdate" }
    if ($EffectiveBranch) {
        $argsList += @("-Branch", $EffectiveBranch)
    }

    Write-Verbose ("Elevation command: powershell.exe " + ($argsList -join " "))

    try {
        Start-Process powershell.exe -ArgumentList $argsList -Verb RunAs | Out-Null
    } catch {
        Write-Info "Elevation failed: $($_.Exception.Message)" ([ConsoleColor]::Red)
        Read-Host "Press Enter to exit"
        break script
    }

    break script
}

function Test-ForAdmin {
    $Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $Run_As_Admin) {
        Write-LogMessage -Message_Type "ERROR" -Message "The script has not been launched with admin rights"
        [System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
        EXIT
    }
    Write-LogMessage -Message_Type "INFO" -Message "The script has been launched with admin rights"
}

function Test-ForSources {
    if (-not (Test-Path -Path $Sources)) {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources folder is missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure.`nThe folder `"Sources`" is missing !!!")
        EXIT
    }
    Write-LogMessage -Message_Type "SUCCESS" -Message "The sources folder exists"
    
    $Check_Sources_Files_Count = (Get-ChildItem -Path "$Current_Folder\Sources\Run_in_Sandbox" -Recurse).count
    if ($Check_Sources_Files_Count -lt 25) {
        Write-LogMessage -Message_Type "ERROR" -Message "Some contents are missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure !!!")
        EXIT
    }
}

function Test-ForSandbox {
    try {
        $Is_Sandbox_Installed = (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue | Where-Object { $_.featurename -eq "Containers-DisposableClientVM" }).state
    } catch {
        if (Test-Path -Path "C:\Windows\System32\WindowsSandbox.exe") {
            Write-LogMessage -Message_Type "WARNING" -Message "It looks like you have the `Windows Sandbox` Feature installed, but your `TrustedInstaller` Service is disabled."
            Write-LogMessage -Message_Type "WARNING" -Message "The Script will continue, but you should check for issues running Windows Sandbox."
            $Is_Sandbox_Installed = "Enabled"
        } else {
            $Is_Sandbox_Installed = "Disabled"
        }
    }
    if ($Is_Sandbox_Installed -eq "Disabled") {
        Write-LogMessage -Message_Type "ERROR" -Message "The feature `Windows Sandbox` is not installed !!!"
        [System.Windows.Forms.MessageBox]::Show("The feature `Windows Sandbox` is not installed !!!")
        EXIT
    }
}

function Test-ForSandboxFolder {
    if ( [string]::IsNullOrEmpty($Sandbox_Folder) ) {
        return
    }
    if (-not (Test-Path -Path $Sandbox_Folder) ) {
        [System.Windows.Forms.MessageBox]::Show("Can not find the folder $Sandbox_Folder")
        EXIT
    }
}

function Copy-Sources {
    try {
        Copy-Item -Path $Sources -Destination $env:ProgramData -Force -Recurse | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "Sources have been copied in $env:ProgramData\Run_in_Sandbox"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources have not been copied in $env:ProgramData\Run_in_Sandbox"
        EXIT
    }
    
    # Copy CommonFunctions.ps1 to the installation directory so RunInSandbox.ps1 can load it
    try {
        Copy-Item -Path "$Current_Folder\CommonFunctions.ps1" -Destination "$env:ProgramData\Run_in_Sandbox\" -Force | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "CommonFunctions.ps1 copied to installation directory"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to copy CommonFunctions.ps1 to installation directory"
        EXIT
    }
    
    if (-not (Test-Path -Path "$env:ProgramData\Run_in_Sandbox\RunInSandbox.ps1") ) {
        Write-LogMessage -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"
        [System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")
        EXIT
    }
}

function Unblock-Sources {
    $Sources_Unblocked = $False
    try {
        Get-ChildItem -Path $Run_in_Sandbox_Folder -Recurse | Unblock-File
        Write-LogMessage -Message_Type "SUCCESS" -Message "Sources files have been unblocked"
        $Sources_Unblocked = $True
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources files have not been unblocked"
        EXIT
    }

    if ($Sources_Unblocked -ne $True) {
        Write-LogMessage -Message_Type "ERROR" -Message "Source files could not be unblocked"
        [System.Windows.Forms.MessageBox]::Show("Source files could not be unblocked")
        EXIT
    }
}

function New-Checkpoint {
    if (-not $NoCheckpoint) {
        $SystemRestoreEnabled = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval").RPSessionInterval
        if ($SystemRestoreEnabled -eq 0) {
            Write-LogMessage -Message_Type "WARNING" -Message "System Restore feature is disabled. Enable this to create a System restore point"
        } else {
            $Checkpoint_Command = '-Command Checkpoint-Computer -Description "Windows_Sandbox_Context_menus" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop'
            $ReturnValue = Start-Process -FilePath "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe" -ArgumentList $Checkpoint_Command -Wait -PassThru -WindowStyle Minimized
            if ($ReturnValue.ExitCode -eq 0) {
                Write-LogMessage -Message_Type "SUCCESS" -Message "Creation of restore point `"Add Windows Sandbox Context menus`""
            } else {
                Write-LogMessage -Message_Type "ERROR" -Message "Creation of restore point `"Add Windows Sandbox Context menus`" failed."
                Write-LogMessage -Message_Type "ERROR" -Message "Press any button to continue anyway."
                Read-Host
            }
        } 
    }
}

Export-ModuleMember -Function @(
    'Test-IsAdmin',
    'Ensure-Admin',
    'Test-ForAdmin',
    'Test-ForSources',
    'Test-ForSandbox',
    'Test-ForSandboxFolder',
    'Copy-Sources',
    'Unblock-Sources',
    'New-Checkpoint'
)