param (
    [Switch]$NoCheckpoint,
    [Switch]$DeepClean,
    [Switch]$AutoUpdate,
    [String]$Branch = ""
)

if (-not $AutoUpdate) {
    $BranchSource = $null

    if ($PSBoundParameters.ContainsKey("Branch") -and $Branch) {
        $BranchSource = "Parameter"
        Write-Host "[DEBUG] Branch parameter supplied: '$Branch'" -ForegroundColor Cyan
    } else {
        Write-Host "[DEBUG] Branch parameter NOT supplied (using detection)" -ForegroundColor Yellow

        if ($args -and $args.Count -gt 0) {
            Write-Host ("[DEBUG] Raw arguments: " + ($args -join ' ')) -ForegroundColor DarkCyan
            for ($i = 0; $i -lt $args.Count; $i++) {
                if ($args[$i] -eq "-Branch" -or $args[$i] -eq "/Branch") {
                    if ($i + 1 -lt $args.Count) {
                        $Branch = $args[$i + 1]
                        $BranchSource = "Args"
                        Write-Host "[DEBUG] Branch extracted from raw arguments: '$Branch'" -ForegroundColor Cyan
                    }
                    break
                }
            }
        }

        if (-not $Branch -and $MyInvocation.Line) {
            Write-Host "[DEBUG] Invocation line: $($MyInvocation.Line)" -ForegroundColor DarkCyan
            if ($MyInvocation.Line -match '(?i)(?:^|\s)-Branch\s+(?:"([^"]+)"|''([^'']+)''|([^\s"`]+))') {
                $BranchCandidate = $Matches[1]
                if (-not $BranchCandidate) { $BranchCandidate = $Matches[2] }
                if (-not $BranchCandidate) { $BranchCandidate = $Matches[3] }
                if ($BranchCandidate) {
                    $Branch = $BranchCandidate
                    $BranchSource = "InvocationLine"
                    Write-Host "[DEBUG] Branch parsed from invocation line: '$Branch'" -ForegroundColor Cyan
                }
            }
        }

        if (-not $Branch) {
            foreach ($__bEnv in @("RIS_BRANCH", "RUN_IN_SANDBOX_BRANCH")) {
                $val = (Get-Item -Path Env:$__bEnv -ErrorAction SilentlyContinue).Value
                if ($val) {
                    $Branch = $val
                    $BranchSource = "EnvVar:$__bEnv"
                    Write-Host "[DEBUG] Branch sourced from environment variable ${__bEnv}: '$Branch'" -ForegroundColor Cyan
                    break
                }
            }
        }
    }

    if ($Branch) {
        Write-Host "[DEBUG] Effective branch detected from $BranchSource => '$Branch'" -ForegroundColor Green
    } else {
        Write-Host "[DEBUG] No branch detected from parameters/args/env (will resolve from version.json or default)" -ForegroundColor Yellow
    }
}

# Resolve environment and branch before any elevation attempts
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$IsInstalled = Test-Path $Run_in_Sandbox_Folder

$ResolvedBranch = $null
if ($Branch) {
    $ResolvedBranch = $Branch
} elseif ($IsInstalled) {
    $VersionJson = Join-Path $Run_in_Sandbox_Folder "version.json"
    if (Test-Path $VersionJson) {
        try {
            $VersionData = Get-Content $VersionJson -Raw | ConvertFrom-Json
            if ($VersionData.branch) {
                $ResolvedBranch = $VersionData.branch
            }
        } catch {
            if (-not $AutoUpdate) {
                Write-Host "[DEBUG] Failed to read version.json: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

if (-not $ResolvedBranch) {
    $ResolvedBranch = "master"
}

$Branch = $ResolvedBranch
if (-not $AutoUpdate) {
    Write-Host "[DEBUG] Branch resolved before elevation: $Branch" -ForegroundColor Cyan
}

# Function to restart the script with admin rights
function Restart-ScriptWithAdmin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $RestartBranch = if ($Branch) { $Branch } else { "master" }

        Write-Host ""
        Write-Host "=== Elevation Required ===" -ForegroundColor Yellow
        Write-Host "[DEBUG] Detected branch: $RestartBranch" -ForegroundColor Cyan

        # Download the installer script for the selected branch into a temporary file
        $TempScript = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "Install_Run-in-Sandbox_Restart.ps1")
        try {
            # Build a tiny wrapper that downloads and executes the installer at elevation, avoiding truncated/cached files
            $InstallerUrl    = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$RestartBranch/Install_Run-in-Sandbox.ps1?_ts=$([DateTime]::UtcNow.Ticks)"
            $NoCheckpointVal = if ($NoCheckpoint) { '$true' } else { '$false' }
            $DeepCleanVal    = if ($DeepClean)    { '$true' } else { '$false' }
            $AutoUpdateVal   = if ($AutoUpdate)   { '$true' } else { '$false' }

            $Wrapper = @"
`$ErrorActionPreference = 'Stop'
`$u = '$InstallerUrl'
`$h = @{ 'User-Agent'='Run-in-Sandbox-Installer/1.0 (+https://github.com/Joly0/Run-in-Sandbox)'; 'Accept'='text/plain' }
`$code = `$null
try { `$code = Invoke-RestMethod -Uri `$u -Headers `$h -UseBasicParsing -TimeoutSec 30 } catch { `$code = `$null }
if (-not `$code -or `$code.TrimStart().StartsWith('<')) {
    `$tmp = [System.IO.Path]::GetTempFileName()
    Invoke-WebRequest -Uri `$u -Headers `$h -UseBasicParsing -OutFile `$tmp -TimeoutSec 30 -ErrorAction Stop
    `$code = Get-Content -Path `$tmp -Raw
    Remove-Item -Path `$tmp -Force -ErrorAction SilentlyContinue
}
if (-not `$code) { throw 'Failed to download installer content' }
`$sb = [ScriptBlock]::Create(`$code)
& `$sb -NoCheckpoint:$NoCheckpointVal -DeepClean:$DeepCleanVal -AutoUpdate:$AutoUpdateVal -Branch '$RestartBranch'
"@

            Set-Content -Path $TempScript -Value $Wrapper -Encoding UTF8
            Write-Host "[DEBUG] Wrote elevated wrapper to: $TempScript" -ForegroundColor Cyan

            # Parse wrapper to ensure validity before elevation
            $null = $tokens = $null; $null = $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($TempScript, [ref]$tokens, [ref]$errors) | Out-Null
            if ($errors -and $errors.Count -gt 0) {
                Write-Host "[ERROR] Elevated wrapper parse errors:" -ForegroundColor Red
                $errors | ForEach-Object { Write-Host ("  " + $_) -ForegroundColor Red }
                Read-Host "Press Enter to exit"
                break script
            }
        } catch {
            Write-Host "[ERROR] Failed to prepare elevated wrapper: $($_.Exception.Message)" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            break script
        }

        # Build the argument list preserving any flags/branch override
        $ArgumentList = @(
            "-NoExit",
            "-ExecutionPolicy", "Bypass",
            "-File", $TempScript
        )

        # Flags are passed inside the EncodedCommand wrapper; do not append any script parameters here.

        Write-Host "[DEBUG] Argument list to elevated PowerShell:" -ForegroundColor Cyan
        $ArgumentList | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkCyan }

        Write-Host "[DEBUG] Command preview:" -ForegroundColor Cyan
        Write-Host ("    powershell.exe " + ($ArgumentList | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' ') -ForegroundColor DarkCyan

        if (-not (Test-Path $TempScript)) {
            Write-Host "[ERROR] Temporary script was not created. Aborting elevation." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            break script
        }

        Write-Host ""
        Read-Host "Press Enter to continue with elevation"

        Start-Process powershell.exe -ArgumentList $ArgumentList -Verb RunAs
        Write-Host "[DEBUG] Elevation request sent." -ForegroundColor Cyan
        Write-Host ""

        break script
    }
}

# Restart the script with admin rights if not already running as admin
Restart-ScriptWithAdmin

# Load latest CommonFunctions once at top-level (online preferred, fallback to local)
try {
    $cf = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$Branch/CommonFunctions.ps1" -UseBasicParsing
    . ([ScriptBlock]::Create($cf))
    if (-not $AutoUpdate) { Write-Host "[DEBUG] CommonFunctions loaded from GitHub ($Branch)" -ForegroundColor Green }
} catch {
    if (Test-Path "$Run_in_Sandbox_Folder\CommonFunctions.ps1") {
        . "$Run_in_Sandbox_Folder\CommonFunctions.ps1"
        if (-not $AutoUpdate) { Write-Host "[WARNING] Using local CommonFunctions fallback" -ForegroundColor Yellow }
    } else {
        throw "CommonFunctions.ps1 could not be loaded"
    }
}
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    break script
}


# At this point $Run_in_Sandbox_Folder, $IsInstalled, and $Branch are already resolved
$BackupCreated = $false

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
        try {
            # Invoke-RestMethod already parses JSON; do not pipe to ConvertFrom-Json
            $LatestData = Invoke-RestMethod -Uri $LatestUrl -UseBasicParsing -TimeoutSec 10
            $LatestVersion = $LatestData.version
        } catch {
            Write-Host "[DEBUG] Failed to retrieve latest version from ${LatestUrl}: $($_.Exception.Message)" -ForegroundColor Yellow
            $LatestVersion = $null
        }
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
                    break script
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
                break script
            }
        }
    }
    
    # Perform deep-clean if requested
    if ($DeepClean) {
        Write-Host ""
        Write-Host "Performing deep-clean..." -ForegroundColor Yellow
        
        
        if (Get-Command Find-RegistryIconPaths -ErrorAction SilentlyContinue) {
            [String[]] $results = @()
            $results = Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT'
            $results += Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT\SystemFileAssociations'
            
            $Current_User_SID = (Get-ChildItem -Path Registry::\HKEY_USERS | Where-Object { Test-Path -Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty -Path "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]
            $HKCU_Classes = "HKEY_USERS\$Current_User_SID" + "_Classes"
            
            $results += Find-RegistryIconPaths -rootRegistryPath $HKCU_Classes
            $results = $results | Where-Object { $_ -notlike "HKEY_CLASSES_ROOT\SystemFileAssociations\SystemFileAssociations*" }
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
        break script
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
    break script
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
    break script
}

# Remove the zip file
try {
    Write-Host "Removing zip file..."
    Remove-Item -Path $zipPath
    Write-Host "Zip file removed."
} catch {
    Write-Error "Failed to remove the zip file: $_"
    break script
}

# Backup config before installation (for merge)
$ConfigBackup = $null
if ($IsInstalled -and (Test-Path "$Run_in_Sandbox_Folder\Sandbox_Config.xml")) {
    $ConfigBackup = "$env:TEMP\Sandbox_Config_Backup.xml"
    Copy-Item "$Run_in_Sandbox_Folder\Sandbox_Config.xml" $ConfigBackup -Force
}

# Construct the path to the add_structure.ps1 script
$addStructureScript = Join-Path -Path $extractPath -ChildPath "Add_Structure.ps1"

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
    
    # Ensure Add_Structure.ps1 runs with its working directory so relative paths (e.g. .\Sources) resolve correctly
    Push-Location $extractPath
    try {
        if ($NoCheckpoint) {
            & ".\Add_Structure.ps1" -NoCheckpoint
        } else {
            & ".\Add_Structure.ps1"
        }
    } finally {
        Pop-Location
    }
    
    # Merge config if it was an update
    if ($ConfigBackup -and (Test-Path $ConfigBackup)) {
        Write-Host "Merging configuration..." -ForegroundColor Cyan
        try {
            Merge-SandboxConfig -OldConfigPath $ConfigBackup -NewConfigPath "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
        } catch {
            # Simple fallback: copy matching element values without XPath
            Write-Host "[WARNING] Merge-SandboxConfig failed, applying simple merge" -ForegroundColor Yellow
            try {
                [xml]$OldCfg = Get-Content $ConfigBackup
                [xml]$NewCfg = Get-Content "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
                if ($OldCfg.Configuration -and $NewCfg.Configuration) {
                    foreach ($child in $OldCfg.Configuration.ChildNodes) {
                        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                        $target = $null
                        foreach ($n in $NewCfg.Configuration.ChildNodes) {
                            if ($n.NodeType -eq [System.Xml.XmlNodeType]::Element -and $n.Name -eq $child.Name) {
                                $target = $n
                                break
                            }
                        }
                        if ($target) {
                            $target.InnerText = $child.InnerText
                        }
                    }
                    $NewCfg.Save("$Run_in_Sandbox_Folder\Sandbox_Config.xml") | Out-Null
                }
            } catch {
                Write-Host "[ERROR] Simple merge failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
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
        break script
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
    break script
}

Write-Host ""
if (-not $AutoUpdate) {
    Read-Host "Press Enter to exit"
}
