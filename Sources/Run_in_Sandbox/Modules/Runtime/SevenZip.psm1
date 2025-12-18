<#
.SYNOPSIS
    7-Zip integration module for Run-in-Sandbox

.DESCRIPTION
    This module provides 7-Zip functionality for the Run-in-Sandbox application.
    It handles archive creation, extraction, and compression operations.
#>

# Function to find 7-Zip installation on host system
function Find-Host7Zip {
    # Try common installation paths
    $CommonPaths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe",
        "${env:ProgramFiles}\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    
    foreach ($Path in $CommonPaths) {
        if (Test-Path $Path) {
            return $Path
        }
    }
    
    # Check registry for installation path
    try {
        $RegPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\7-Zip" -Name "Path" -ErrorAction SilentlyContinue
        if ($RegPath -and (Test-Path "$($RegPath.Path)\7z.exe")) {
            return "$($RegPath.Path)\7z.exe"
        }
    } catch {}
    
    # Check PATH environment variable
    try {
        $7zInPath = Get-Command "7z.exe" -ErrorAction SilentlyContinue
        if ($7zInPath) {
            return $7zInPath.Source
        }
    } catch {}
    
    return $null
}

# Function to get latest 7-Zip download URL from GitHub releases
function Get-Latest7ZipDownloadUrl {
    try {
        $ApiUrl = "https://api.github.com/repos/ip7z/7zip/releases/latest"
        $Response = Invoke-RestMethod -Uri $ApiUrl -UseBasicParsing
        
        # Look for x64 MSI installer first, fallback to x86 if needed
        $Asset = $Response.assets | Where-Object { $_.name -like "*-x64.msi" -and $_.name -notlike "*extra*" }
        
        if (-not $Asset) {
            # Fallback to x86 MSI if x64 not available
            $Asset = $Response.assets | Where-Object { $_.name -like "*.msi" -and $_.name -notlike "*extra*" -and $_.name -notlike "*x64*" }
        }
        
        if ($Asset) {
            return $Asset.browser_download_url
        }
    } catch {
        Write-LogMessage -Message_Type "WARNING" -Message "Failed to get latest 7-Zip version from GitHub: $($_.Exception.Message)"
    }
    
    return $null
}


# Function to check if cached 7-Zip installer should be updated
function Test-7ZipCacheAge {
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    $CachedInstaller = "$TempFolder\7zSetup.msi"
    $VersionFile = "$TempFolder\7zVersion.txt"
    
    # If no cached installer exists, we need to download
    if (-not (Test-Path $CachedInstaller)) {
        return $true
    }
    
    # Check if cache is older than 7 days
    $CacheAge = (Get-Date) - (Get-Item $CachedInstaller).LastWriteTime
    if ($CacheAge.Days -gt 7) {
        Write-LogMessage -Message_Type "INFO" -Message "Cached 7-Zip installer is $($CacheAge.Days) days old, checking for updates"
        return $true
    }
    
    return $false
}

# Function to download and cache latest 7-Zip installer
function Update-7ZipCache {
    param(
        [switch]$Force
    )
    
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    $CachedInstaller = "$TempFolder\7zSetup.msi"
    $VersionFile = "$TempFolder\7zVersion.txt"
    
    # Check if we need to update (unless forced)
    if (-not $Force -and -not (Test-7ZipCacheAge)) {
        Write-LogMessage -Message_Type "INFO" -Message "Cached 7-Zip installer is recent, skipping update"
        return $true
    }
    
    # Ensure temp folder exists
    if (-not (Test-Path $TempFolder)) {
        New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
    }
    
    # Get latest download URL
    $DownloadUrl = Get-Latest7ZipDownloadUrl
    if (-not $DownloadUrl) {
        Write-LogMessage -Message_Type "ERROR" -Message "Could not determine latest 7-Zip download URL"
        return $false
    }
    
    try {
        Write-LogMessage -Message_Type "INFO" -Message "Downloading latest 7-Zip installer from: $DownloadUrl"
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $CachedInstaller -UseBasicParsing
        
        # Save download timestamp and URL for tracking
        @{
            Downloaded = (Get-Date).ToString()
            Url = $DownloadUrl
        } | ConvertTo-Json | Set-Content $VersionFile
        
        Write-LogMessage -Message_Type "SUCCESS" -Message "7-Zip installer cached successfully"
        return $true
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to download 7-Zip installer: $($_.Exception.Message)"
        return $false
    }
}

# Function to initialize 7-Zip cache - ensures it's available and current
function Initialize-7ZipCache {
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $TempFolder = "$Run_in_Sandbox_Folder\temp"
    $CachedInstaller = "$TempFolder\7zSetup.msi"
    
    # Try to update if needed (network available)
    try {
        if (Test-7ZipCacheAge) {
            Update-7ZipCache
        }
    } catch {
        Write-LogMessage -Message_Type "WARNING" -Message "Could not check for 7-Zip updates, using cached version if available"
    }
    
    # Return whether we have a usable cached installer
    return (Test-Path $CachedInstaller)
}

Export-ModuleMember -Function @(
    'Find-Host7Zip',
    'Get-Latest7ZipDownloadUrl',
    'Test-7ZipCacheAge',
    'Update-7ZipCache',
    'Initialize-7ZipCache'
)
