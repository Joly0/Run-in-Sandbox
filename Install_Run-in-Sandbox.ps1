#Requires -Version 5.1
<#
.SYNOPSIS
    Installer script for Run-in-Sandbox

.DESCRIPTION
    Downloads and installs Run-in-Sandbox from GitHub. Supports branch selection,
    deep-clean for legacy installations, and backup creation before updates.

.PARAMETER Branch
    The GitHub branch to install from. Defaults to 'master' for new installations
    or the currently installed branch for updates.

.PARAMETER DeepClean
    Performs a deep-clean of legacy registry entries before installation.
    This removes old context menu entries and takes 5-10 minutes.

.PARAMETER NoCheckpoint
    Skips creation of a system restore point during installation.

.PARAMETER RepoOwner
    GitHub owner of the repository to install from (e.g. a fork). Defaults to
    'Joly0' for new installations or the owner stored in version.json for updates.

.PARAMETER RepoName
    GitHub repository name to install from. Defaults to 'Run-in-Sandbox' for new
    installations or the name stored in version.json for updates.

.PARAMETER OriginalUserSid
    Internal use only. Carries the SID of the non-elevated user across the UAC
    boundary so HKCU writes target the original user (not the admin account
    that answered the UAC prompt). Set automatically by Invoke-AsAdmin.

.EXAMPLE
    irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1 | iex

.EXAMPLE
    .\Install_Run-in-Sandbox.ps1 -Branch dev -DeepClean

.EXAMPLE
    .\Install_Run-in-Sandbox.ps1 -RepoOwner myuser -RepoName Run-in-Sandbox
#>
[CmdletBinding()]
param (
    [switch]$NoCheckpoint,
    [switch]$DeepClean,
    [string]$Branch,
    [string]$RepoOwner,
    [string]$RepoName,
    [string]$OriginalUserSid
)

# Resolve the original user's SID before any elevation happens. If we were
# relaunched by Invoke-AsAdmin, the parent already captured it and forwarded it
# to us. Otherwise, capture the current token now - but only if it represents a
# real user account, since service SIDs would later defeat the explorer.exe
# fallback in Resolve-InteractiveUserSid.
if ($OriginalUserSid) {
    $Global:OriginalUserSid = $OriginalUserSid
} else {
    try {
        $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        if ($sid -notin @('S-1-5-18','S-1-5-19','S-1-5-20') -and $sid -notlike 'S-1-5-80-*') {
            $Global:OriginalUserSid = $sid
        }
    } catch { }
}

if ($VerbosePreference -eq 'Continue') {
    $PSDefaultParameterValues['*:Verbose'] = $true
} else {
    $PSDefaultParameterValues.Remove('*:Verbose') | Out-Null
}

# ======================================================================================
# Configuration
# ======================================================================================
$DefaultRepoOwner = "Joly0"
$DefaultRepoName = "Run-in-Sandbox"

# Globals
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$IsInstalled = Test-Path $Run_in_Sandbox_Folder

# Without -Branch/-RepoOwner/-RepoName, updates stay on the installed branch and
# repository (version.json), new installs use master of Joly0/Run-in-Sandbox.
# Resolved here already because the modules below are downloaded from there.
$InstalledVersionData = $null
$InstalledVersionJson = Join-Path $Run_in_Sandbox_Folder "version.json"
if (Test-Path $InstalledVersionJson) {
    try {
        $InstalledVersionData = Get-Content $InstalledVersionJson -Raw | ConvertFrom-Json
    } catch { }
}
if (-not $Branch) {
    $Branch = if ($InstalledVersionData.branch) { $InstalledVersionData.branch } else { "master" }
}
if (-not $RepoOwner) {
    $RepoOwner = if ($InstalledVersionData.repoOwner) { $InstalledVersionData.repoOwner } else { $DefaultRepoOwner }
}
if (-not $RepoName) {
    $RepoName = if ($InstalledVersionData.repoName) { $InstalledVersionData.repoName } else { $DefaultRepoName }
}
# Read by the modules (Environment.psm1 falls back to Joly0/Run-in-Sandbox)
$Global:Repo_Owner = $RepoOwner
$Global:Repo_Name = $RepoName
Write-Verbose "Repository: $RepoOwner/$RepoName"

# ======================================================================================
# Function to dynamically load modules from GitHub
# ======================================================================================
function Import-ModuleFromGitHub {
    param(
        [string]$ModulePath,
        [string]$BranchName
    )
    
    $moduleUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$BranchName/$ModulePath"
    try {
        Write-Verbose "Loading module from: $moduleUrl"
        
        # Create a temporary directory for modules if it doesn't exist
        $tempModulePath = Join-Path $env:TEMP "RunInSandboxModules"
        if (-not (Test-Path $tempModulePath)) {
            New-Item -Path $tempModulePath -ItemType Directory -Force | Out-Null
        }
        
        # Create the full path for the temporary module file
        $moduleName = Split-Path $ModulePath -Leaf
        $tempModuleFile = Join-Path $tempModulePath $moduleName
        
        # Download the module content to a temporary file
        $moduleContent = Invoke-RestMethod -Uri $moduleUrl -UseBasicParsing -TimeoutSec 30
        Set-Content -Path $tempModuleFile -Value $moduleContent -Force
        
        # Import the module using the proper Import-Module cmdlet
        Import-Module $tempModuleFile -Force -Global
        
        # Clean up the temporary file
        Remove-Item $tempModuleFile -Force -ErrorAction SilentlyContinue
        
        return $true
    } catch {
        Write-Verbose "Failed to load module $ModulePath`: $($_.Exception.Message)"
        return $false
    }
}


# ======================================================================================
# Load required modules from GitHub
# ======================================================================================
$requiredModules = @(
    "Sources/Run_in_Sandbox/Modules/Shared/Logging.psm1",
    "Sources/Run_in_Sandbox/Modules/Shared/Version.psm1",
    "Sources/Run_in_Sandbox/Modules/Shared/Environment.psm1",
    "Sources/Run_in_Sandbox/Modules/Shared/Config.psm1",
    "Sources/Run_in_Sandbox/Modules/Installer/Core.psm1",
    "Sources/Run_in_Sandbox/Modules/Installer/Registry.psm1",
    "Sources/Run_in_Sandbox/Modules/Installer/Validation.psm1"
)
foreach ($modulePath in $requiredModules) {
    if (-not (Import-ModuleFromGitHub -ModulePath $modulePath -BranchName $Branch)) {
        # No fallback: CommonFunctions.ps1 lacks most installer functions, and
        # the package itself is downloaded from GitHub further below anyway
        Write-Host "Failed to load $modulePath from $RepoOwner/$RepoName (branch '$Branch')." -ForegroundColor Red
        Write-Host "Check your internet connection and that the repository and branch exist. Run with -Verbose for details." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        break script
    }
}

# ======================================================================================
# Branch resolution and elevation
# ======================================================================================
$Branch = Resolve-Branch -Requested $Branch -Installed:$IsInstalled
Write-Verbose "Effective branch: $Branch"

Invoke-AsAdmin -EffectiveBranch $Branch -NoCheckpoint:$NoCheckpoint -DeepClean:$DeepClean -OriginalUserSid $Global:OriginalUserSid

# ======================================================================================
# Show banner if existing installation detected
# ======================================================================================
if ($IsInstalled) {
    Write-Info "Run-in-Sandbox detected." ([ConsoleColor]::Cyan)
}

# ======================================================================================
# Version info and optional reinstall prompt
# ======================================================================================
$BackupCreated = $false
if ($IsInstalled) {
    $CurrentVersion = Get-CurrentVersionSimple
    $LatestVersion  = Get-LatestVersionFromBranch -EffectiveBranch $Branch
    $InstalledBranch = Get-InstalledBranch

    if ($CurrentVersion) {
        Write-Info ("Current Version:   {0}" -f $CurrentVersion) ([ConsoleColor]::Green)
        $branchToShow = if ($InstalledBranch) { $InstalledBranch } else { $Branch }
        Write-Info ("Current Branch:    {0}" -f $branchToShow) ([ConsoleColor]::Cyan)
        if ($LatestVersion) {
            Write-Info ("Latest Version:    {0}" -f $LatestVersion) ([ConsoleColor]::Green)
        }
        if ($LatestVersion -and $Branch) {
            Write-Info ("Requested Branch:  {0}" -f $Branch) ([ConsoleColor]::Green)
        }
        Write-Host ""

        if ($LatestVersion -and $CurrentVersion -match '^\d{4}-\d{2}-\d{2}' -and $LatestVersion -match '^\d{4}-\d{2}-\d{2}') {
            $currentDate = [datetime]::ParseExact($CurrentVersion, 'yyyy-MM-dd', $null)
            $latestDate  = [datetime]::ParseExact($LatestVersion,  'yyyy-MM-dd', $null)
            if ($latestDate -le $currentDate) {
                Write-Info "You are already running the latest version." ([ConsoleColor]::Green)
                Write-Host ""
                $userResponse = Read-Host "Do you want to reinstall anyway? (Y/N)"
                if ($userResponse -notmatch '^(?i)y') {
                    Write-Info "Installation cancelled." ([ConsoleColor]::Yellow)
                    Read-Host "Press Enter to exit"
                    break script
                }
            }
        }
    }

    Write-Info ("This will UPDATE your existing installation (Branch: {0})" -f $Branch) ([ConsoleColor]::Yellow)
    Write-Host ""

    # Backup
    $BackupCreated = New-InstallBackup
    if (-not $BackupCreated) {
        $userResponse = Read-Host "Continue without backup? (Y/N)"
        if ($userResponse -notmatch '^(?i)y') {
            Write-Info "Update cancelled." ([ConsoleColor]::Yellow)
            Read-Host "Press Enter to exit"
            break script
        }
    }

    # Optional DeepClean prompt
    $DeepClean = Get-DeepCleanConsent -DeepCleanRef:$DeepClean
    Invoke-DeepCleanIfRequested -DeepClean:$DeepClean

    Write-Host ""
    Write-Info "Proceeding with update installation..." ([ConsoleColor]::Cyan)
    Write-Host ""
}


# ======================================================================================
# Download, Extract, Install
# ======================================================================================
$extractPath = $null
try {
    $extractPath = Install-PackageArchive -EffectiveBranch $Branch
} catch {
    Write-Info $_ ([ConsoleColor]::Red)
    if ($BackupCreated) {
        Write-Info "Backup available at: $Run_in_Sandbox_Folder\backup" ([ConsoleColor]::Yellow)
    }
    break script
}

# Preserve config for merge (if update)
$DefaultStartupNames = Get-DefaultStartupScriptNames -ExtractPath $extractPath
Remove-OldInstallIfDeepClean -DeepClean:$DeepClean -RunFolder:$Run_in_Sandbox_Folder -DefaultNames $DefaultStartupNames

# If we performed a deep-clean, treat this run as a fresh install for the rest of the flow
if ($DeepClean) { $IsInstalled = $false }

$ConfigBackup = $null
if ($IsInstalled -and (Test-Path "$Run_in_Sandbox_Folder\Sandbox_Config.xml")) {
    $ConfigBackup = "$env:TEMP\Sandbox_Config_Backup.xml"
    Copy-Item "$Run_in_Sandbox_Folder\Sandbox_Config.xml" $ConfigBackup -Force
}

try {
    Invoke-AddStructure -ExtractPath $extractPath -NoCheckpoint:$NoCheckpoint -IsInstalled:$IsInstalled
    
    # Restore config backup after Add_Structure (which overwrites everything)
    # This ensures user settings are preserved before Update-CoreFiles merges them
    if ($ConfigBackup -and (Test-Path $ConfigBackup)) {
        $destConfig = Join-Path $Run_in_Sandbox_Folder "Sandbox_Config.xml"
        Copy-Item $ConfigBackup $destConfig -Force
        Write-Verbose "Restored Sandbox_Config.xml from backup after Add_Structure"
    }
    
    # Sync files for both updates and fresh installations
    # Update-CoreFiles handles config merging for updates and fresh copy for new installs
    $syncParams = @{
        ExtractPath = $extractPath
        RunFolder = $Run_in_Sandbox_Folder
        DefaultNames = $DefaultStartupNames
        IsInstalled = $IsInstalled
    }
    Update-CoreFiles @syncParams
    
    # Not tied to $IsInstalled: a deep-clean resets that flag, but is exactly the
    # case where custom startup scripts were backed up (no-op otherwise)
    Restore-CustomStartupScripts -RunFolder $Run_in_Sandbox_Folder
    Get-VersionJson -RunFolder $Run_in_Sandbox_Folder -ExtractPath $extractPath -EffectiveBranch $Branch -LatestVersion $LatestVersion

    $valid = Test-Installation -RunFolder $Run_in_Sandbox_Folder
    if (-not $valid) {
        Write-Info "Installation validation failed!" ([ConsoleColor]::Red)
        if ($BackupCreated) {
            Write-Info "Backup available at: $Run_in_Sandbox_Folder\backup" ([ConsoleColor]::Yellow)
        }
        break script
    }

    Clear-TempArtifacts -ExtractPath $extractPath -RunFolder $Run_in_Sandbox_Folder
    
    # Clean up config backup
    if ($ConfigBackup -and (Test-Path $ConfigBackup)) {
        Remove-Item $ConfigBackup -Force -ErrorAction SilentlyContinue
    }

    if ($IsInstalled) {
        Write-Info ("Update completed successfully. Branch: {0}" -f $Branch) ([ConsoleColor]::Green)
    } else {
        Write-Info "Installation completed successfully." ([ConsoleColor]::Green)
    }
} catch {
    Write-Info ("Failed to execute installation: {0}" -f $_) ([ConsoleColor]::Red)
    if ($BackupCreated) {
        Write-Info "Backup available at: $Run_in_Sandbox_Folder\backup" ([ConsoleColor]::Yellow)
        Write-Info "To restore, copy contents from backup back to $Run_in_Sandbox_Folder" ([ConsoleColor]::Yellow)
    }
    break script
}

Write-Host ""
$PSDefaultParameterValues.Remove('*:Verbose') | Out-Null
Read-Host "Press Enter to exit"
