param (
    [Switch]$NoCheckpoint,
    [Switch]$DeepClean
)

# Function to restart the script with admin rights
function Restart-ScriptWithAdmin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $params = "-NoProfile -ExecutionPolicy Bypass -NoExit -Command `"(Invoke-webrequest -URI `"https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1`").Content | Invoke-Expression"
        if ($NoCheckpoint) { $params += " -NoCheckpoint" }
        if ($DeepClean) { $params += " -DeepClean" }
        $params += "`""
        Start-Process powershell.exe $params -Verb RunAs
        exit
    }
}

# Restart the script with admin rights if not already running as admin
Restart-ScriptWithAdmin

# Define paths
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$IsInstalled = Test-Path $Run_in_Sandbox_Folder

# Check if already installed and offer update
if ($IsInstalled) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Run-in-Sandbox Detected" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Load CommonFunctions to check version
    if (Test-Path "$Run_in_Sandbox_Folder\CommonFunctions.ps1") {
        . "$Run_in_Sandbox_Folder\CommonFunctions.ps1"
        
        # Get version info
        $VersionInfo = Get-VersionInfo
        if ($VersionInfo.Current -and $VersionInfo.Latest) {
            Write-Host "Current Version: " -NoNewline -ForegroundColor Gray
            Write-Host $VersionInfo.Current -ForegroundColor Green
            Write-Host "Latest Version:  " -NoNewline -ForegroundColor Gray
            Write-Host $VersionInfo.Latest -ForegroundColor Green
            Write-Host ""
            
            try {
                $CurrentDate = [DateTime]::ParseExact($VersionInfo.Current, 'yyyy-MM-dd', $null)
                $LatestDate = [DateTime]::ParseExact($VersionInfo.Latest, 'yyyy-MM-dd', $null)
                
                if ($LatestDate -le $CurrentDate) {
                    Write-Host "You are already running the latest version!" -ForegroundColor Green
                    Write-Host ""
                    $Response = Read-Host "Do you want to reinstall anyway? (Y/N)"
                    if ($Response -ne 'Y' -and $Response -ne 'y') {
                        Write-Host "Installation cancelled." -ForegroundColor Yellow
                        Read-Host "Press Enter to exit"
                        exit
                    }
                }
            } catch {}
        }
    }
    
    Write-Host "This will UPDATE your existing installation." -ForegroundColor Yellow
    Write-Host ""
    
    # Offer deep-clean option
    if (-not $DeepClean) {
        Write-Host "Deep-Clean Option:" -ForegroundColor Cyan
        Write-Host "  A deep-clean removes ALL registry entries with Run-in-Sandbox icon paths." -ForegroundColor Gray
        Write-Host "  This ensures a completely fresh installation but takes longer." -ForegroundColor Gray
        Write-Host "  Recommended if you had previous installation issues." -ForegroundColor Gray
        Write-Host ""
        $DeepCleanResponse = Read-Host "Perform deep-clean before update? (Y/N)"
        if ($DeepCleanResponse -eq 'Y' -or $DeepCleanResponse -eq 'y') {
            $DeepClean = $true
        }
    }
    
    # Perform deep-clean if requested
    if ($DeepClean) {
        Write-Host ""
        Write-Host "Performing deep-clean..." -ForegroundColor Yellow
        
        # Load CommonFunctions if not already loaded
        if (-not (Get-Command Find-RegistryIconPaths -ErrorAction SilentlyContinue)) {
            if (Test-Path "$Run_in_Sandbox_Folder\CommonFunctions.ps1") {
                . "$Run_in_Sandbox_Folder\CommonFunctions.ps1"
            }
        }
        
        if (Get-Command Find-RegistryIconPaths -ErrorAction SilentlyContinue) {
            [String[]] $results = @()
            $results = Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT'
            $results += Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT\SystemFileAssociations'
            
            $Current_User_SID = (Get-ChildItem -Path Registry::\HKEY_USERS | Where-Object { Test-Path -Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty -Path "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]
            $HKCU_Classes = "Registry::HKEY_USERS\$Current_User_SID" + "_Classes"
            
            $results += Find-RegistryIconPaths -rootRegistryPath $HKCU_Classes
            $results = $results | Where-Object { $_ -notlike "REGISTRY::HKEY_CLASSES_ROOT\SystemFileAssociations\SystemFileAssociations*" }
            $results = $results | Select-Object -Unique | Sort-Object
            
            foreach ($reg_path in $results) {
                try {
                    Get-ChildItem -Path $reg_path -Recurse | Sort-Object { $_.PSPath.Split('\').Count } -Descending | Select-Object -Property PSPath -ExpandProperty PSPath | Remove-Item -Force -Confirm:$false -ErrorAction Stop
                    if (Test-Path -Path $reg_path) {
                        Remove-Item -LiteralPath $reg_path -Force -Recurse -Confirm:$false -ErrorAction Stop
                    }
                    Write-Host "  ✓ Removed: $reg_path" -ForegroundColor Green
                } catch {
                    Write-Host "  ✗ Failed to remove: $reg_path" -ForegroundColor Red
                }
            }
            Write-Host "Deep-clean completed." -ForegroundColor Green
        } else {
            Write-Host "Warning: Deep-clean functions not available, skipping..." -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
    Write-Host "Proceeding with update installation..." -ForegroundColor Cyan
    Write-Host ""
}

# Define the URL and file paths
$zipUrl = "https://github.com/Joly0/Run-in-Sandbox/archive/refs/heads/master.zip"
$tempPath = [System.IO.Path]::GetTempPath()
$zipPath = Join-Path -Path $tempPath -ChildPath "master.zip"
$extractPath = Join-Path -Path $tempPath -ChildPath "Run-in-Sandbox-master"

# Remove existing extracted folder if it exists
if (Test-Path $extractPath) {
    try {
        Write-Host "Removing existing extracted folder..."
        Remove-Item -Path $extractPath -Recurse -Force
        Write-Host "Existing extracted folder removed."
    } catch {
        Write-Error "Failed to remove existing extracted folder: $_"
        exit 1
    }
}

# Download the zip file
try {
    Write-Host "Downloading zip file..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    $ProgressPreference = 'Continue'
    Write-Host "Download completed."
} catch {
    Write-Error "Failed to download the zip file: $_"
    exit 1
}

# Extract the zip file
try {
    Write-Host "Extracting zip file..."
    $ProgressPreference = 'SilentlyContinue'
    Expand-Archive -Path $zipPath -DestinationPath $tempPath -Force
    $ProgressPreference = 'Continue'
    Write-Host "Extraction completed."
} catch {
    Write-Error "Failed to extract the zip file: $_"
    exit 1
}

# Remove the zip file
try {
    Write-Host "Removing zip file..."
    Remove-Item -Path $zipPath
    Write-Host "Zip file removed."
} catch {
    Write-Error "Failed to remove the zip file: $_"
    exit 1
}

# Construct the path to the add_structure.ps1 script
$addStructureScript = Join-Path -Path $extractPath -ChildPath "Add_Structure.ps1"

# Set Execution Policy and unblock files
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted
Get-ChildItem -LiteralPath $extractPath -Filter "*.ps1" | Unblock-File

# Execute the add_structure.ps1 script with parameters
try {
    if ($IsInstalled) {
        Write-Host "Updating Run-in-Sandbox..." -ForegroundColor Cyan
    } else {
        Write-Host "Installing Run-in-Sandbox..." -ForegroundColor Cyan
    }
    
    if ($NoCheckpoint) {
        & $addStructureScript -NoCheckpoint
    } else {
        & $addStructureScript
    }
    
    if ($IsInstalled) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  Update completed successfully!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  Installation completed successfully!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to execute add_structure.ps1: $_"
    exit 1
}

Write-Host ""
Read-Host "Press Enter to exit"
