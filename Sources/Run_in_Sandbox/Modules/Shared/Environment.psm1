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

# Resolves the SID of the interactive user that should be targeted with HKCU writes.
# Tries three sources, in order:
#   1. $Global:OriginalUserSid - set by the parent installer before elevation,
#      so a UAC prompt answered with a different admin account does not flip
#      HKCU to that admin's hive
#   2. The owner of explorer.exe, when either:
#      - the current token is a service SID (Intune SYSTEM context), or
#      - we are elevated as a different account than the explorer owner (the
#        UAC-with-different-creds case for users who run Add_Structure.ps1
#        directly, where there is no wrapper installer to forward the SID)
#      Only applied when exactly one distinct owner is found - multi-user
#      Terminal Server sessions are genuinely ambiguous and fall through.
#   3. The current process token (own creds, or already-correct elevation)
function Resolve-InteractiveUserSid {
    $ServiceSids = @('S-1-5-18','S-1-5-19','S-1-5-20')

    if ($Global:OriginalUserSid) {
        return $Global:OriginalUserSid
    }

    $WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $CurrentSid = $WindowsIdentity.User.Value
    $IsServiceContext = $CurrentSid -in $ServiceSids -or $CurrentSid -like 'S-1-5-80-*'
    # Inline elevation check - Test-IsAdmin is defined later in this module
    $IsElevated = ([Security.Principal.WindowsPrincipal]::new($WindowsIdentity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    try {
        $ExplorerOwners = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop |
            ForEach-Object {
                $r = Invoke-CimMethod -InputObject $_ -MethodName GetOwnerSid -ErrorAction SilentlyContinue
                if ($r -and $r.ReturnValue -eq 0) { $r.Sid }
            } | Sort-Object -Unique)
        if ($ExplorerOwners.Count -eq 1) {
            $InteractiveSid = $ExplorerOwners[0]
            if ($IsServiceContext -or ($IsElevated -and $InteractiveSid -ne $CurrentSid)) {
                return $InteractiveSid
            }
        }
    } catch { }

    # No reliable interactive user found - fall back to the current token
    return $CurrentSid
}

# Current user SID and registry paths
$Global:CurrentSid = Resolve-InteractiveUserSid
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
        [switch]$DeepClean,
        [string]$OriginalUserSid
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
    # Forward the pre-elevation user SID so the elevated process targets the
    # original user's HKCU instead of the admin account that answered UAC.
    if ($OriginalUserSid) {
        $argsList += @("-OriginalUserSid", $OriginalUserSid)
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

# Sets Modify permissions for the built-in Users group on the given path so that
# non-admin users can edit config files and startup scripts at runtime.
# Uses the well-known SID S-1-5-32-545 (BUILTIN\Users) which works regardless of
# the OS display language.
function Set-UserWritePermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [switch]$IsDirectory
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Verbose "Set-UserWritePermissions: Path does not exist, skipping: $Path"
        return $false
    }

    try {
        $acl = Get-Acl -LiteralPath $Path
        # BUILTIN\Users – language-independent SID
        $identity = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-545")

        if ($IsDirectory) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity,
                "Modify",
                "ContainerInherit,ObjectInherit",
                "None",
                "Allow"
            )
        } else {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity,
                "Modify",
                "None",
                "None",
                "Allow"
            )
        }

        $acl.SetAccessRule($rule)
        Set-Acl -LiteralPath $Path -AclObject $acl
        Write-Verbose "Set-UserWritePermissions: Permissions set on $Path"
        return $true
    } catch {
        Write-Verbose "Set-UserWritePermissions: Failed on ${Path}: $($_.Exception.Message)"
        return $false
    }
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
    'Invoke-AsAdmin',
    'Set-UserWritePermissions',
    'Resolve-InteractiveUserSid'
)
