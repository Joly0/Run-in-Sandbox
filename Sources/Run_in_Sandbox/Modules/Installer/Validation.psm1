<#
.SYNOPSIS
    Validation and setup module for Run-in-Sandbox installation.

.DESCRIPTION
    This module provides functions for validating system requirements, checking
    prerequisites, and performing setup operations during Run-in-Sandbox installation.
    Includes admin privilege checks, source file validation, Windows Sandbox feature
    detection, hardware prerequisite checks, and file operations.
#>

# Import shared modules for logging
$ModuleRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$SharedModulePath = Join-Path $ModuleRoot "Modules\Shared"
if (Test-Path "$SharedModulePath\Logging.psm1") {
    Import-Module "$SharedModulePath\Logging.psm1" -Force -Global
}

# Load Windows Forms assembly for message boxes
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# Module-level variables
$script:Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

# Checks if the script is running with administrator privileges
function Test-ForAdmin {
    $Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $Run_As_Admin) {
        Write-LogMessage -Message_Type "ERROR" -Message "The script has not been launched with admin rights"
        [System.Windows.Forms.MessageBox]::Show("Please run the tool with admin rights :-)")
        throw "The script has not been launched with admin rights"
    }
    Write-LogMessage -Message_Type "INFO" -Message "The script has been launched with admin rights"
}


# Checks if the Sources folder exists and contains required files
function Test-ForSources {
    param (
        [Parameter(Mandatory=$true)] [string]$Current_Folder,
        [Parameter(Mandatory=$true)] [string]$Sources
    )
    
    if (-not (Test-Path -Path $Sources)) {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources folder is missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure.`nThe folder `"Sources`" is missing !!!")
        throw "Sources folder is missing"
    }
    Write-LogMessage -Message_Type "SUCCESS" -Message "The sources folder exists"
    
    $Check_Sources_Files_Count = (Get-ChildItem -Path "$Current_Folder\Sources\Run_in_Sandbox" -Recurse).count
    if ($Check_Sources_Files_Count -lt 25) {  # Reduced from 40 to 26 (removed 14 bundled 7zip files)
        Write-LogMessage -Message_Type "ERROR" -Message "Some contents are missing"
        [System.Windows.Forms.MessageBox]::Show("It seems you haven´t downloaded all the folder structure !!!")
        throw "Some source contents are missing"
    }
}

# Checks if the Windows Sandbox feature is installed and enabled.
# If missing, offers to enable it with user confirmation (requires manual reboot).
function Test-ForSandbox {
    $sandboxState = $null

    try {
        $sandboxState = (Get-WindowsOptionalFeature -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.FeatureName -eq "Containers-DisposableClientVM" }).State
    } catch {
        # If Get-WindowsOptionalFeature fails (e.g. TrustedInstaller disabled), fall back to exe check
        if (Test-Path -Path "C:\Windows\System32\WindowsSandbox.exe") {
            Write-LogMessage -Message_Type "WARNING" -Message "Could not query optional features (TrustedInstaller may be disabled), but WindowsSandbox.exe exists."
            Write-LogMessage -Message_Type "WARNING" -Message "The script will continue, but you should verify Windows Sandbox works correctly."
            return
        }
        $sandboxState = "Disabled"
    }

    if ($sandboxState -eq "Enabled") {
        Write-LogMessage -Message_Type "SUCCESS" -Message "Windows Sandbox is enabled"
        return
    }

    # Sandbox is not enabled — offer to enable it
    Write-LogMessage -Message_Type "WARNING" -Message "Windows Sandbox feature is not enabled"
    Write-Host ""
    $response = Read-Host "Would you like to enable Windows Sandbox now? A manual reboot will be required afterwards. (Y/N)"

    if ($response -notmatch '^(?i)y') {
        Write-LogMessage -Message_Type "ERROR" -Message "Windows Sandbox is not enabled. Installation cancelled by user."
        [System.Windows.Forms.MessageBox]::Show("Windows Sandbox is not enabled.`nPlease enable it manually and run the installer again.")
        throw "Windows Sandbox is not enabled. Installation cancelled by user."
    }

    # Enable the Sandbox feature
    Write-LogMessage -Message_Type "INFO" -Message "Enabling Windows Sandbox feature (Containers-DisposableClientVM)..."
    try {
        Enable-WindowsOptionalFeature -Online -FeatureName "Containers-DisposableClientVM" -NoRestart -All -ErrorAction Stop | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "Windows Sandbox feature has been enabled"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to enable Windows Sandbox: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to enable Windows Sandbox.`nPlease enable it manually and run the installer again.")
        throw "Failed to enable Windows Sandbox: $($_.Exception.Message)"
    }

    # Feature enabled — tell the user to reboot and re-run
    Write-LogMessage -Message_Type "INFO" -Message "Please reboot your computer, then run the installer again."
    [System.Windows.Forms.MessageBox]::Show("Windows Sandbox has been enabled successfully.`nPlease reboot your computer and run the installer again.")
    throw "Reboot required after enabling Windows Sandbox. Please restart and re-run the installer."
}

# Checks hardware and system prerequisites (RAM, disk space, virtualization)
function Test-Prerequisites {
    # RAM check - Windows Sandbox requires at least 4 GB
    $ramGB = [math]::Round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    if ($ramGB -lt 4) {
        Write-LogMessage -Message_Type "ERROR" -Message "Not enough RAM: $ramGB GB found, at least 4 GB required"
        [System.Windows.Forms.MessageBox]::Show("Not enough RAM: $ramGB GB found, at least 4 GB required.")
        throw "Not enough RAM: $ramGB GB found, at least 4 GB required"
    }
    Write-LogMessage -Message_Type "SUCCESS" -Message "RAM check passed: $ramGB GB available"

    # Disk space check - need at least 1 GB free on system drive
    $diskFreeGB = [math]::Round((Get-PSDrive -Name C).Free / 1GB, 2)
    if ($diskFreeGB -lt 1) {
        Write-LogMessage -Message_Type "ERROR" -Message "Not enough free disk space: $diskFreeGB GB found, at least 1 GB required"
        [System.Windows.Forms.MessageBox]::Show("Not enough free disk space: $diskFreeGB GB found, at least 1 GB required.")
        throw "Not enough free disk space: $diskFreeGB GB found, at least 1 GB required"
    }
    Write-LogMessage -Message_Type "SUCCESS" -Message "Disk space check passed: $diskFreeGB GB free"
}

# Checks if the specified Sandbox folder exists
function Test-ForSandboxFolder {
    param (
        [string]$Sandbox_Folder
    )
    
    if ( [string]::IsNullOrEmpty($Sandbox_Folder) ) {
        return
    }
    if (-not (Test-Path -Path $Sandbox_Folder) ) {
        [System.Windows.Forms.MessageBox]::Show("Can not find the folder $Sandbox_Folder")
        throw "Can not find the folder $Sandbox_Folder"
    }
}


# Copies source files to the installation directory
function Copy-Sources {
    param (
        [Parameter(Mandatory=$true)] [string]$Current_Folder,
        [Parameter(Mandatory=$true)] [string]$Sources
    )
    
    try {
        Copy-Item -Path $Sources -Destination $env:ProgramData -Force -Recurse | Out-Null
        Write-LogMessage -Message_Type "SUCCESS" -Message "Sources have been copied in $env:ProgramData\Run_in_Sandbox"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources have not been copied in $env:ProgramData\Run_in_Sandbox"
        throw "Sources have not been copied in $env:ProgramData\Run_in_Sandbox: $_"
    }
    
    if (-not (Test-Path -Path "$env:ProgramData\Run_in_Sandbox\RunInSandbox.ps1") ) {
        Write-LogMessage -Message_Type "ERROR" -Message "File RunInSandbox.ps1 is missing"
        [System.Windows.Forms.MessageBox]::Show("File RunInSandbox.ps1 is missing !!!")
        throw "File RunInSandbox.ps1 is missing after copy"
    }
}

# Unblocks all source files in the installation directory
function Unblock-Sources {
    $Sources_Unblocked = $False
    try {
        Get-ChildItem -Path $script:Run_in_Sandbox_Folder -Recurse | Unblock-File
        Write-LogMessage -Message_Type "SUCCESS" -Message "Sources files have been unblocked"
        $Sources_Unblocked = $True
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Sources files have not been unblocked"
        throw "Sources files have not been unblocked: $_"
    }

    if ($Sources_Unblocked -ne $True) {
        Write-LogMessage -Message_Type "ERROR" -Message "Source files could not be unblocked"
        [System.Windows.Forms.MessageBox]::Show("Source files could not be unblocked")
        throw "Source files could not be unblocked"
    }
}

# Creates a system restore checkpoint before making changes
function New-Checkpoint {
    param (
        [switch]$NoCheckpoint
    )
    
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

# Export public functions
Export-ModuleMember -Function @(
    'Test-ForAdmin',
    'Test-ForSources',
    'Test-ForSandbox',
    'Test-ForSandboxFolder',
    'Test-Prerequisites',
    'Copy-Sources',
    'Unblock-Sources',
    'New-Checkpoint'
)
