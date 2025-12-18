<#
.SYNOPSIS
    Version management module for Run-in-Sandbox

.DESCRIPTION
    This module provides version management functionality for the Run-in-Sandbox application.
    It handles branch resolution, version detection, and version file management.
    This is the basic version for installation purposes (excludes auto-update notification functions).
#>

# Global variable for Run-in-Sandbox folder
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

# Helper function for colored console output
function Write-Info {
    param(
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::White
    )
    Write-Host $Message -ForegroundColor $Color
}

# Resolves which branch to use for installation
function Resolve-Branch {
    [CmdletBinding()]
    param(
        [string]$Requested,
        [bool]$Installed
    )
    # Precedence:
    # 1) Explicit -Branch parameter
    # 2) version.json (if already installed)
    # 3) default: master
    if ($Requested) {
        Write-Verbose "Branch via parameter: '$Requested'"
        return $Requested
    }

    if ($Installed) {
        $versionJsonPath = Join-Path $Run_in_Sandbox_Folder "version.json"
        if (Test-Path $versionJsonPath) {
            try {
                $versionData = Get-Content $versionJsonPath -Raw | ConvertFrom-Json
                if ($versionData.branch) {
                    Write-Verbose "Branch via existing version.json = '$($versionData.branch)'"
                    return $versionData.branch
                }
            } catch {
                Write-Verbose "Could not read branch from version.json: $($_.Exception.Message)"
            }
        }
    }

    Write-Verbose "No branch provided/detected - defaulting to 'master'"
    return "master"
}

# Gets the current installed version from version.json or config
function Get-CurrentVersionSimple {
    # Lightweight: version.json preferred; fallback to config
    $currentVersion = $null
    $versionJsonPath = Join-Path $Run_in_Sandbox_Folder "version.json"
    if (Test-Path $versionJsonPath) {
        try {
            $versionData = Get-Content $versionJsonPath -Raw | ConvertFrom-Json
            $currentVersion = $versionData.version
        } catch { }
    }
    if (-not $currentVersion -and (Test-Path "$Run_in_Sandbox_Folder\Sandbox_Config.xml")) {
        try {
            [xml]$configXml = Get-Content "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
            $currentVersion = $configXml.Configuration.CurrentVersion
        } catch { }
    }
    return $currentVersion
}


# Gets the currently installed branch from version.json
function Get-InstalledBranch {
    $versionJsonPath = Join-Path $Run_in_Sandbox_Folder "version.json"
    if (Test-Path $versionJsonPath) {
        try {
            $versionData = Get-Content $versionJsonPath -Raw | ConvertFrom-Json
            if ($versionData.branch) { return $versionData.branch }
        } catch { }
    }
    return $null
}

# Fetches the latest version from the specified branch on GitHub
function Get-LatestVersionFromBranch {
    [CmdletBinding()]
    param([string]$EffectiveBranch)
    try {
        $url = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$EffectiveBranch/version.json"
        $data = Invoke-RestMethod -Uri $url -UseBasicParsing -TimeoutSec 15
        return $data.version
    } catch {
        Write-Verbose "Latest version fetch failed: $($_.Exception.Message)"
        return $null
    }
}

# Creates or updates the version.json file in the installation folder
function Get-VersionJson {
    param(
        [string]$RunFolder,
        [string]$ExtractPath,
        [string]$EffectiveBranch,
        [string]$LatestVersion
    )

    $extractedVersion = $null
    $versionJsonInExtract = Join-Path $ExtractPath "version.json"
    if (Test-Path $versionJsonInExtract) {
        try {
            $versionData = Get-Content $versionJsonInExtract -Raw | ConvertFrom-Json
            $extractedVersion = $versionData.version
        } catch { }
    }

    if (-not $extractedVersion -and $LatestVersion) {
        $extractedVersion = $LatestVersion
    }
    if (-not $extractedVersion) {
        $extractedVersion = (Get-Date).ToString("yyyy-MM-dd")
    }

    @{ version = $extractedVersion; branch = $EffectiveBranch } |
        ConvertTo-Json |
        Set-Content (Join-Path $RunFolder "version.json")
}

# Tests if the installation is valid by checking for required files
function Test-Installation {
    param([string]$RunFolder)
    
    $requiredFiles = @(
        "RunInSandbox.ps1",
        "Sandbox_Config.xml"
    )
    
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $RunFolder $file
        if (-not (Test-Path $filePath)) {
            Write-Verbose "Required file missing: $filePath"
            return $false
        }
    }
    
    # Check for Modules folder
    $modulesPath = Join-Path $RunFolder "Modules"
    if (-not (Test-Path $modulesPath)) {
        Write-Verbose "Modules folder missing: $modulesPath"
        return $false
    }
    
    return $true
}

Export-ModuleMember -Function @(
    'Write-Info',
    'Resolve-Branch',
    'Get-CurrentVersionSimple',
    'Get-InstalledBranch',
    'Get-LatestVersionFromBranch',
    'Get-VersionJson',
    'Test-Installation'
)
