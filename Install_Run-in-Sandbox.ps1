#Requires -Version 5.1
[CmdletBinding()]
param (
    [switch]$NoCheckpoint,
    [switch]$DeepClean,
    [switch]$AutoUpdate,
    [string]$Branch = "master"
)

if ($VerbosePreference -eq 'Continue') {
    $PSDefaultParameterValues['*:Verbose'] = $true
} else {
    $PSDefaultParameterValues.Remove('*:Verbose') | Out-Null
}

# ======================================================================================
# Minimal, maintainable installer with functions and optional verbose logging
# - Keep functionality intact
# - Reduce console noise (use Write-Verbose for debug)
# - Simplify elevation and download flow
# - Dynamically loads modules from GitHub repository
# ======================================================================================

# Configuration
$DefaultRepoOwner = "Joly0"
$RepoName = "Run-in-Sandbox"

# Globals
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$IsInstalled = Test-Path $Run_in_Sandbox_Folder

# Function to dynamically load modules from GitHub
function Import-ModuleFromGitHub {
    param(
        [string]$ModulePath,
        [string]$Branch = $DefaultBranch,
        [string]$RepoOwner = $DefaultRepoOwner
    )
    
    $moduleUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$ModulePath"
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
        Import-Module $tempModuleFile -Force
        
        # Clean up the temporary file
        Remove-Item $tempModuleFile -Force -ErrorAction SilentlyContinue
        
        return $true
    } catch {
        Write-Verbose "Failed to load module $ModulePath`: $($_.Exception.Message)"
        return $false
    }
}

# Load required modules from GitHub
$moduleLoadSuccess = $true
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Shared/Logging.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Shared/Version.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Shared/Environment.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Shared/Config.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Installer/Core.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Installer/Registry.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)
$moduleLoadSuccess = $moduleLoadSuccess -and (Import-ModuleFromGitHub -ModulePath "Sources/Run_in_Sandbox/Modules/Installer/Validation.psm1" -Branch $Branch -RepoOwner $DefaultRepoOwner)

if (-not $moduleLoadSuccess) {
    Write-Info "Failed to load modules from GitHub. This might be due to network issues or an invalid branch." ([ConsoleColor]::Red)
    Write-Info "Falling back to CommonFunctions.ps1..." ([ConsoleColor]::Yellow)
    
    # Fallback to CommonFunctions.ps1
    try {
        $commonFunctionsUrl = "https://raw.githubusercontent.com/$DefaultRepoOwner/$RepoName/$Branch/CommonFunctions.ps1"
        Write-Verbose ("Loading CommonFunctions from: {0}" -f $commonFunctionsUrl)
        $commonFunctionsContent = Invoke-RestMethod -Uri $commonFunctionsUrl -UseBasicParsing -TimeoutSec 45
        . ([ScriptBlock]::Create($commonFunctionsContent)) # dot-source into script scope
        Write-Verbose "CommonFunctions loaded from GitHub."
    } catch {
        $localCommonFunctionsPath = Join-Path $Run_in_Sandbox_Folder "CommonFunctions.ps1"
        if (Test-Path $localCommonFunctionsPath) {
            . $localCommonFunctionsPath
            Write-Verbose "CommonFunctions loaded from local path."
        } else {
            throw "CommonFunctions.ps1 could not be loaded."
        }
    }
}

# -------------------------------------------------------------------------------------------------
# Branch resolution and elevation
# -------------------------------------------------------------------------------------------------
$Branch = Resolve-Branch -Requested $Branch -Installed:$IsInstalled
Write-Verbose "Effective branch: $Branch"

Invoke-AsAdmin -EffectiveBranch $Branch -NoCheckpoint:$NoCheckpoint -DeepClean:$DeepClean -AutoUpdate:$AutoUpdate

# -------------------------------------------------------------------------------------------------
# Show minimal banner if not AutoUpdate
# -------------------------------------------------------------------------------------------------
if (-not $AutoUpdate -and $IsInstalled) {
    Write-Info "Run-in-Sandbox detected." ([ConsoleColor]::Cyan)
}

# -------------------------------------------------------------------------------------------------
# Version info and optional reinstall prompt (only when not AutoUpdate)
# -------------------------------------------------------------------------------------------------
$BackupCreated = $false
if ($IsInstalled) {
    $CurrentVersion = Get-CurrentVersionSimple
    $LatestVersion  = Get-LatestVersionFromBranch -EffectiveBranch $Branch
    $InstalledBranch = Get-InstalledBranch

    if (-not $AutoUpdate -and $CurrentVersion) {
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

        if ($LatestVersion -and $CurrentVersion -match '^\d{4}-\d{2}-\d{2}$' -and $LatestVersion -match '^\d{4}-\d{2}-\d{2}$') {
            $currentDate = [datetime]::ParseExact($CurrentVersion, 'yyyy-MM-dd', $null)
            $latestDate  = [datetime]::ParseExact($LatestVersion,  'yyyy-MM-dd', $null)
            if ($latestDate -le $currentDate) {
                Write-Info "You are already running the latest version." ([ConsoleColor]::Green)
                Write-Host ""
                $userResponse = Read-Host "Do you want to reinstall anyway? (Y/N)"
                if ($userResponse -notmatch '^(?i)y$') {
                    Write-Info "Installation cancelled." ([ConsoleColor]::Yellow)
                    Read-Host "Press Enter to exit"
                    break script
                }
            }
        }
    }

    if (-not $AutoUpdate) {
        Write-Info ("This will UPDATE your existing installation (Branch: {0})" -f $Branch) ([ConsoleColor]::Yellow)
        Write-Host ""
    }

    # Backup
    $BackupCreated = New-InstallBackup
    if (-not $BackupCreated -and -not $AutoUpdate) {
        $userResponse = Read-Host "Continue without backup? (Y/N)"
        if ($userResponse -notmatch '^(?i)y$') {
            Write-Info "Update cancelled." ([ConsoleColor]::Yellow)
            Read-Host "Press Enter to exit"
            break script
        }
    }

    # Optional DeepClean prompt (only interactive)
    $DeepClean = Get-DeepCleanConsent -AutoUpdate:$AutoUpdate -DeepCleanRef:$DeepClean
    Invoke-DeepCleanIfRequested -DeepClean:$DeepClean

    if (-not $AutoUpdate) {
        Write-Host ""
        Write-Info "Proceeding with update installation..." ([ConsoleColor]::Cyan)
        Write-Host ""
    }
}

# -------------------------------------------------------------------------------------------------
# Download, Extract, Install
# -------------------------------------------------------------------------------------------------
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
    Merge-ConfigIfNeeded -IsInstalled:$IsInstalled -RunFolder:$Run_in_Sandbox_Folder
    
    # Only sync files for updates, not for new installations
    if ($IsInstalled) {
        $syncParams = @{
            ExtractPath = $extractPath
            RunFolder = $Run_in_Sandbox_Folder
            DefaultNames = $DefaultStartupNames
        }
        Update-CoreFiles @syncParams
        Restore-CustomStartupScripts -RunFolder $Run_in_Sandbox_Folder
    }
    Get-VersionJson -RunFolder $Run_in_Sandbox_Folder -ExtractPath $extractPath -EffectiveBranch $Branch -LatestVersion $LatestVersion

    $valid = Test-Installation -RunFolder $Run_in_Sandbox_Folder
    if (-not $valid) {
        Write-Info "Installation validation failed!" ([ConsoleColor]::Red)
        if ($BackupCreated) {
            Write-Info "Backup available at: $Run_in_Sandbox_Folder\backup" ([ConsoleColor]::Yellow)
        }
        break script
    }

    # Set permissions to make the folder writable by everyone
    try {
        Write-Info "Setting folder permissions..." ([ConsoleColor]::Cyan)
        $acl = Get-Acl $Run_in_Sandbox_Folder
        
        # Try multiple approaches for different Windows versions
        $identity = $null
        try {
            # First try using well-known SID for Users
            $identity = [System.Security.Principal.SecurityIdentifier]("S-1-5-32-545")
        } catch {
            try {
                # Fallback to built-in users group
                $identity = [System.Security.Principal.NTAccount]("BUILTIN\Users")
            } catch {
                try {
                    # Last resort - try Everyone
                    $identity = [System.Security.Principal.NTAccount]("Everyone")
                } catch {
                    Write-Info "Warning: Could not create identity for permissions." ([ConsoleColor]::Yellow)
                }
            }
        }
        
        if ($identity) {
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity,
                "FullControl",
                "ContainerInherit,ObjectInherit",
                "None",
                "Allow"
            )
            $acl.SetAccessRule($accessRule)
            Set-Acl $Run_in_Sandbox_Folder $acl
            Write-Info "Folder permissions set successfully." ([ConsoleColor]::Green)
        }
    } catch {
        Write-Info "Warning: Failed to set folder permissions: $($_.Exception.Message)" ([ConsoleColor]::Yellow)
    }

    Clear-TempArtifacts -ExtractPath $extractPath -RunFolder $Run_in_Sandbox_Folder

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
if (-not $AutoUpdate) {
    $PSDefaultParameterValues.Remove('*:Verbose') | Out-Null
    Read-Host "Press Enter to exit"
}
