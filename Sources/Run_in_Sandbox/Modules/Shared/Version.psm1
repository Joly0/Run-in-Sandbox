<#
.SYNOPSIS
    Version management module for Run-in-Sandbox

.DESCRIPTION
    This module provides version management functionality for the Run-in-Sandbox application.
    It handles branch resolution, version detection, and version file management.
#>

# Global variable for Run-in-Sandbox folder
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

function Resolve-Branch {
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

function Get-LatestVersionFromBranch {
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

function Ensure-VersionJson {
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

Export-ModuleMember -Function @(
    'Resolve-Branch',
    'Get-CurrentVersionSimple',
    'Get-InstalledBranch',
    'Get-LatestVersionFromBranch',
    'Ensure-VersionJson'
)