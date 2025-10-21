param (
    [Switch]$NoCheckpoint,
    [Switch]$DeepClean,
    [Switch]$AutoUpdate,
    [String]$Branch = ""
)

# Function to restart the script with admin rights
function Restart-ScriptWithAdmin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $RestartBranch = if ($Branch) { $Branch } else { "master" }
        
        # Build parameter string
        $scriptParams = @()
        if ($NoCheckpoint) { $scriptParams += "-NoCheckpoint" }
        if ($DeepClean) { $scriptParams += "-DeepClean" }
        if ($AutoUpdate) { $scriptParams += "-AutoUpdate" }
        if ($Branch) { $scriptParams += "-Branch"; $scriptParams += $Branch }
        
        # Simple approach: download script, create scriptblock, execute with parameters
        $cmd = "`$script = irm 'https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$RestartBranch/Install_Run-in-Sandbox.ps1'; & ([scriptblock]::Create(`$script)) $($scriptParams -join ' ')"
        
        Start-Process powershell.exe -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-NoExit","-Command",$cmd -Verb RunAs
        exit
    }
}

# Restart the script with admin rights if not already running as admin
Restart-ScriptWithAdmin

# Define paths
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$IsInstalled = Test-Path $Run_in_Sandbox_Folder
$BackupCreated = $false

# Determine branch to use
if (-not $Branch) {
    # Detect from existing installation
    if ($IsInstalled) {
        $VersionJson = "$Run_in_Sandbox_Folder\version.json"
        if (Test-Path $VersionJson) {
            $VersionData = Get-Content $VersionJson -Raw | ConvertFrom-Json
            $Branch = if ($VersionData.branch) { $VersionData.branch } else { "master" }
        } else {
            $Branch = "master"
        }
    } else {
        $Branch = "master"
    }
}

# Check if already installed and show version info
if ($IsInstalled) {
    if (-not $AutoUpdate) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  Run-in-Sandbox Detected" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
    }
    
    # Get current version (simple - no function calls)
    $CurrentVersion = $null
    $VersionJson = "$Run_in_Sandbox_Folder\version.json"
    if (Test-Path $VersionJson) {
        $VersionData = Get-Content $VersionJson -Raw | ConvertFrom-Json
        $CurrentVersion = $VersionData.version
    } elseif (Test-Path "$Run_in_Sandbox_Folder\Sandbox_Config.xml") {
        [xml]$Config = Get-Content "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
        $CurrentVersion = $Config.Configuration.CurrentVersion
    }
    
    # Get latest version
    $LatestVersion = $null
    if (-not $AutoUpdate) {
        $LatestUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$Branch/version.json"
        $LatestData = Invoke-RestMethod -Uri $LatestUrl -UseBasicParsing | ConvertFrom-Json
        $LatestVersion = $LatestData.version
    }
    
    # Display version info
    if (-not $AutoUpdate -and $CurrentVersion) {
        Write-Host "Current Version: " -NoNewline -ForegroundColor Gray
        Write-Host $CurrentVersion -ForegroundColor Green
        Write-Host "Current Branch:  " -NoNewline -ForegroundColor Gray
        Write-Host $Branch -ForegroundColor Cyan
        if ($LatestVersion) {
            Write-Host "Latest Version:  " -NoNewline -ForegroundColor Gray
            Write-Host $LatestVersion -ForegroundColor Green
        }
        Write-Host ""
        
        # Check if already latest
        if ($LatestVersion -and $CurrentVersion -match '^\d{4}-\d{2}-\d{2}$' -and $LatestVersion -match '^\d{4}-\d{2}-\d{2}$') {
            $CurrentDate = [DateTime]::ParseExact($CurrentVersion, 'yyyy-MM-dd', $null)
            $LatestDate = [DateTime]::ParseExact($LatestVersion, 'yyyy-MM-dd', $null)
            
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
        }
    }
    
    if (-not $AutoUpdate) {
        Write-Host "This will UPDATE your existing installation." -ForegroundColor Yellow
        Write-Host "Branch: $Branch" -ForegroundColor Cyan
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
    }
    
    # Create backup before update
    try {
        Write-Host "Creating backup..." -ForegroundColor Yellow
        $BackupFolder = "$Run_in_Sandbox_Folder\backup"
        if (Test-Path $BackupFolder) { Remove-Item -Path $BackupFolder -Recurse -Force }
        New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
        
        Get-ChildItem -Path $Run_in_Sandbox_Folder | Where-Object { $_.Name -notin @("temp", "backup", "logs") } | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination (Join-Path $BackupFolder $_.Name) -Recurse -Force
        }
        
        $BackupCreated = $true
        Write-Host "Backup created successfully." -ForegroundColor Green
    } catch {
        Write-Host "Warning: Backup creation failed: $_" -ForegroundColor Yellow
        if (-not $AutoUpdate) {
            $Response = Read-Host "Continue without backup? (Y/N)"
            if ($Response -ne 'Y' -and $Response -ne 'y') {
                Write-Host "Update cancelled." -ForegroundColor Yellow
                Read-Host "Press Enter to exit"
                exit 1
            }
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
    
    if (-not $AutoUpdate) {
        Write-Host ""
        Write-Host "Proceeding with update installation..." -ForegroundColor Cyan
        Write-Host ""
    }
}

# Define the URL and file paths (branch-aware)
$zipUrl = "https://github.com/Joly0/Run-in-Sandbox/archive/refs/heads/$Branch.zip"
$tempPath = [System.IO.Path]::GetTempPath()
$zipPath = Join-Path -Path $tempPath -ChildPath "Run-in-Sandbox-$Branch.zip"
$extractPath = Join-Path -Path $tempPath -ChildPath "Run-in-Sandbox-$Branch"

# Remove existing extracted folder if it exists
if (Test-Path $extractPath) {
    try {
        Write-Host "Removing existing extracted folder..."
        Remove-Item -Path $extractPath -Recurse -Force
        Write-Host "Existing extracted folder removed."
    } catch {
        Write-Error "Failed to remove existing extracted folder: $_"
        if ($BackupCreated) {
            Write-Host "Attempting rollback..." -ForegroundColor Yellow
            # Rollback logic here if needed
        }
        exit 1
    }
}

# Download the zip file
try {
    Write-Host "Downloading from branch '$Branch'..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    $ProgressPreference = 'Continue'
    Write-Host "Download completed."
} catch {
    Write-Error "Failed to download the zip file: $_"
    if ($BackupCreated) {
        Write-Host "Backup available at: $Run_in_Sandbox_Folder\backup" -ForegroundColor Yellow
    }
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

# Backup config before installation (for merge)
$ConfigBackup = $null
if ($IsInstalled -and (Test-Path "$Run_in_Sandbox_Folder\Sandbox_Config.xml")) {
    $ConfigBackup = "$env:TEMP\Sandbox_Config_Backup.xml"
    Copy-Item "$Run_in_Sandbox_Folder\Sandbox_Config.xml" $ConfigBackup -Force
}

# Construct the path to the add_structure.ps1 script
$addStructureScript = Join-Path -Path $extractPath -ChildPath "Run-in-Sandbox-$Branch\Add_Structure.ps1"

# Set Execution Policy and unblock files
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted
Get-ChildItem -LiteralPath $extractPath -Recurse -Filter "*.ps1" | Unblock-File

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
    
    # Merge config if it was an update
    if ($ConfigBackup -and (Test-Path $ConfigBackup)) {
        Write-Host "Merging configuration..." -ForegroundColor Cyan
        . "$Run_in_Sandbox_Folder\CommonFunctions.ps1"
        Merge-SandboxConfig -OldConfigPath $ConfigBackup -NewConfigPath "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
        Remove-Item $ConfigBackup -Force
    }
    
    # Validate installation
    $ValidationPassed = $true
    $RequiredFiles = @(
        "$Run_in_Sandbox_Folder\Sources\Run_in_Sandbox\RunInSandbox.ps1",
        "$Run_in_Sandbox_Folder\CommonFunctions.ps1",
        "$Run_in_Sandbox_Folder\Sandbox_Config.xml",
        "$Run_in_Sandbox_Folder\version.json"
    )
    
    foreach ($File in $RequiredFiles) {
        if (-not (Test-Path $File)) {
            Write-Host "Validation failed: Missing $File" -ForegroundColor Red
            $ValidationPassed = $false
        }
    }
    
    if (-not $ValidationPassed) {
        Write-Host "Installation validation failed!" -ForegroundColor Red
        if ($BackupCreated) {
            Write-Host "Backup available at: $Run_in_Sandbox_Folder\backup" -ForegroundColor Yellow
        }
        exit 1
    }
    
    # Clean up temp files
    if (Test-Path $extractPath) {
        Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    # Clear update state if it exists
    $StateFile = "$Run_in_Sandbox_Folder\temp\UpdateState.json"
    if (Test-Path $StateFile) { Remove-Item $StateFile -Force }
    
    if ($IsInstalled) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  Update completed successfully!" -ForegroundColor Green
        Write-Host "  Branch: $Branch" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "  Installation completed successfully!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to execute add_structure.ps1: $_"
    if ($BackupCreated) {
        Write-Host "Backup available at: $Run_in_Sandbox_Folder\backup" -ForegroundColor Yellow
        Write-Host "To restore, copy contents from backup folder back to $Run_in_Sandbox_Folder" -ForegroundColor Yellow
    }
    exit 1
}

Write-Host ""
if (-not $AutoUpdate) {
    Read-Host "Press Enter to exit"
}
