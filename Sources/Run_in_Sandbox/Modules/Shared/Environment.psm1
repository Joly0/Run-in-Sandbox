<#
.SYNOPSIS
    Environment module for Run-in-Sandbox

.DESCRIPTION
    This module provides global environment variables for the Run-in-Sandbox application.
    It initializes paths, registry locations, and other environment-specific settings
    needed by installer and runtime scripts.
#>

# Global variables for paths
$Global:Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$Global:Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
$Global:XML_Config = "$Global:Run_in_Sandbox_Folder\Sandbox_Config.xml"

# Logging variables
$Global:TEMP_Folder = $env:temp
$Global:Log_File = "$Global:TEMP_Folder\RunInSandbox_Install.log"

# Windows version detection
$Global:Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption

# Current user SID and registry paths
# Use a more reliable method to get the current user's SID
$Global:CurrentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
if (Test-Path -LiteralPath "Registry::HKEY_USERS\$Global:CurrentSid\Volatile Environment" -ErrorAction SilentlyContinue) {
    $Global:Current_User_SID = $Global:CurrentSid
    $Global:HKCU = "Registry::HKEY_USERS\$Global:Current_User_SID"
    $Global:HKCU_Classes = "Registry::HKEY_USERS\${Global:Current_User_SID}_Classes"
} else {
    # Fallback for cases where the SID-based path doesn't exist
    $Global:HKCU = 'HKCU:'
    $Global:HKCU_Classes = 'HKCU:\Software\Classes'
}

# Load Windows Forms assembly for message boxes (used by validation functions)
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# ======================================================================================
# Installer Helper Functions
# ======================================================================================

# Checks if the current process is running with administrator privileges
function Test-IsAdmin {
    $windowsIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $windowsPrincipal = [Security.Principal.WindowsPrincipal]::new($windowsIdentity)
    return $windowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Restarts the installer script with administrator privileges if not already elevated
function Invoke-AsAdmin {
    [CmdletBinding()]
    param(
        [string]$EffectiveBranch,
        [switch]$NoCheckpoint,
        [switch]$DeepClean
    )
    if (Test-IsAdmin) {
        Write-Verbose "Already elevated."
        return
    }

    Write-Host "Elevation required. Restarting installer as Administrator..." -ForegroundColor Yellow

    # Simplified, reliable elevation: download current branch installer to temp and re-run with same parameters.
    $TempScript = Join-Path ([IO.Path]::GetTempPath()) "Install_Run-in-Sandbox.Elevated.ps1"
    $InstallerUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$EffectiveBranch/Install_Run-in-Sandbox.ps1"

    try {
        Write-Verbose "Downloading elevated installer from: $InstallerUrl"
        Invoke-WebRequest -Uri $InstallerUrl -UseBasicParsing -OutFile $TempScript -TimeoutSec 45 -ErrorAction Stop
    } catch {
        Write-Host "Failed to download installer for elevation: $($_.Exception.Message)" -ForegroundColor Red
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
    if ($EffectiveBranch) {
        $argsList += @("-Branch", $EffectiveBranch)
    }

    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')) {
        $argsList += "-Verbose"
        Write-Verbose "Forwarding -Verbose to elevated process."
    } else {
        Write-Verbose "Not forwarding -Verbose (caller did not supply it)."
    }

    Write-Verbose ("Elevation command: powershell.exe " + ($argsList -join " "))

    try {
        Start-Process powershell.exe -ArgumentList $argsList -Verb RunAs | Out-Null
    } catch {
        Write-Host "Elevation failed: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        break script
    }

    break script
}

# Export variables and functions
Export-ModuleMember -Variable @(
    'Run_in_Sandbox_Folder',
    'Sandbox_Icon',
    'XML_Config',
    'TEMP_Folder',
    'Log_File',
    'Windows_Version',
    'Current_User_SID',
    'HKCU',
    'HKCU_Classes'
) -Function @(
    'Test-IsAdmin',
    'Invoke-AsAdmin'
)
