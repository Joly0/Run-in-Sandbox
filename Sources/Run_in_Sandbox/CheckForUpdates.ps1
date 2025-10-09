#Requires -Version 5.1

<#
.SYNOPSIS
    Manually check for Run-in-Sandbox updates
    
.DESCRIPTION
    This script checks GitHub for the latest Run-in-Sandbox version
    and displays available updates with changelog information.
    
.PARAMETER ShowDialog
    Skip console output and go directly to changelog dialog if update available
    
.EXAMPLE
    .\CheckForUpdates.ps1
    Interactive console check with option to view changelog
    
.EXAMPLE
    .\CheckForUpdates.ps1 -ShowDialog
    Directly show dialog if update available
    
.NOTES
    Version: 1.0
    Author: Run-in-Sandbox Project
#>

param(
    [switch]$ShowDialog
)

# Set paths
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

# Load common functions
if (-not (Test-Path "$Run_in_Sandbox_Folder\CommonFunctions.ps1")) {
    Write-Host "Error: Run-in-Sandbox not properly installed" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

. "$Run_in_Sandbox_Folder\CommonFunctions.ps1"

# Console banner
if (-not $ShowDialog) {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Run-in-Sandbox Update Checker" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}

# Get version information
$VersionInfo = Get-VersionInfo

if (-not $VersionInfo.Current) {
    Write-Host "Error: Could not determine current version" -ForegroundColor Red
    Write-Host "Please reinstall Run-in-Sandbox" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not $ShowDialog) {
    Write-Host "Current Version: " -NoNewline -ForegroundColor Gray
    Write-Host $VersionInfo.Current -ForegroundColor Green
    Write-Host "Checking for updates..." -ForegroundColor Yellow
    Write-Host ""
}

if (-not $VersionInfo.Latest) {
    Write-Host "Error: Could not check for updates" -ForegroundColor Red
    Write-Host "Please check your internet connection" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not $ShowDialog) {
    Write-Host "Latest Version:  " -NoNewline -ForegroundColor Gray
    Write-Host $VersionInfo.Latest -ForegroundColor Green
    Write-Host ""
}

# Compare versions
try {
    $CurrentDate = [DateTime]::ParseExact($VersionInfo.Current, 'yyyy-MM-dd', $null)
    $LatestDate = [DateTime]::ParseExact($VersionInfo.Latest, 'yyyy-MM-dd', $null)
    $Comparison = $CurrentDate.CompareTo($LatestDate)
} catch {
    Write-Host "Error: Invalid version format" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

if ($Comparison -lt 0) {
    # Update available
    if (-not $ShowDialog) {
        Write-Host "┌─────────────────────────────────────────┐" -ForegroundColor Green
        Write-Host "│  UPDATE AVAILABLE!                      │" -ForegroundColor Green  
        Write-Host "└─────────────────────────────────────────┘" -ForegroundColor Green
        Write-Host ""
        Write-Host "A new version is available: " -NoNewline -ForegroundColor Cyan
        Write-Host $VersionInfo.Latest -ForegroundColor White
        Write-Host ""
        
        $Response = Read-Host "View changelog? (Y/N)"
        
        if ($Response -eq 'Y' -or $Response -eq 'y') {
            Show-ChangelogDialog -LatestVersion $VersionInfo.Latest
        }
    }
    else {
        # Direct to dialog
        Show-ChangelogDialog -LatestVersion $VersionInfo.Latest
    }
}
elseif ($Comparison -eq 0) {
    # Already up to date
    if (-not $ShowDialog) {
        Write-Host "✓ You are already running the latest version!" -ForegroundColor Green
        Write-Host ""
        Read-Host "Press Enter to exit"
    }
}
else {
    # Current version is newer (development version)
    if (-not $ShowDialog) {
        Write-Host "⚠ You are running a development version" -ForegroundColor Yellow
        Write-Host "  Current: $($VersionInfo.Current)" -ForegroundColor Yellow
        Write-Host "  Latest stable: $($VersionInfo.Latest)" -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to exit"
    }
}