<#
.SYNOPSIS
    Core installer module for Run-in-Sandbox

.DESCRIPTION
    This module provides core installation functionality for the Run-in-Sandbox application.
    It handles the main installation process and coordination of installation steps.
#>

function Request-OptionalDeepClean {
    param([switch]$AutoUpdate, [switch]$DeepCleanRef)
    if ($AutoUpdate) { return $DeepCleanRef }

    if (-not $DeepCleanRef) {
        Write-Info "Deep-Clean removes legacy registry icon path entries for a fully clean install." ([ConsoleColor]::Cyan)
        $deepCleanAnswer = Read-Host "Perform deep-clean before update? (Y/N)"
        if ($deepCleanAnswer -match '^(?i)y$') { return $true }
    }
    return $DeepCleanRef
}

function New-InstallBackup {
    try {
        Write-Info "Creating backup..." ([ConsoleColor]::Yellow)
        $BackupFolder = Join-Path $Run_in_Sandbox_Folder "backup"
        if (Test-Path $BackupFolder) {
            Remove-Item -Path $BackupFolder -Recurse -Force
        }
        New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null

        Get-ChildItem -Path $Run_in_Sandbox_Folder | Where-Object { $_.Name -notin @("temp", "backup", "logs") } | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination (Join-Path $BackupFolder $_.Name) -Recurse -Force
        }
        Write-Info "Backup created successfully." ([ConsoleColor]::Green)
        return $true
    } catch {
        Write-Info "Warning: Backup creation failed: $($_.Exception.Message)" ([ConsoleColor]::Yellow)
        return $false
    }
}

function Invoke-DeepCleanIfRequested {
    param([switch]$DeepClean)
    if (-not $DeepClean) { return }
    Write-Info "Performing deep-clean..." ([ConsoleColor]::Yellow)

    if (Get-Command Find-RegistryIconPaths -ErrorAction SilentlyContinue) {
        [string[]]$registryPaths = @()
        $registryPaths = Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT'
        $registryPaths += Find-RegistryIconPaths -rootRegistryPath 'HKEY_CLASSES_ROOT\SystemFileAssociations'

        $currentUserSid = (Get-ChildItem -Path Registry::\HKEY_USERS | Where-Object { Test-Path -Path "$($_.pspath)\Volatile Environment" } | ForEach-Object { (Get-ItemProperty -Path "$($_.pspath)\Volatile Environment") }).PSParentPath.split("\")[-1]
        $hkcuClassesPath = "HKEY_USERS\$currentUserSid" + "_Classes"

        $registryPaths += Find-RegistryIconPaths -rootRegistryPath $hkcuClassesPath
        $registryPaths = $registryPaths | Where-Object { $_ -notlike "HKEY_CLASSES_ROOT\SystemFileAssociations\SystemFileAssociations*" }
        $registryPaths = $registryPaths | Select-Object -Unique | Sort-Object

        foreach ($registryPath in $registryPaths) {
            try {
                Get-ChildItem -Path $registryPath -Recurse |
                    Sort-Object { $_.PSPath.Split('\').Count } -Descending |
                    Select-Object -ExpandProperty PSPath |
                    Remove-Item -Force -Confirm:$false -ErrorAction Stop

                if (Test-Path -Path $registryPath) {
                    Remove-Item -LiteralPath $registryPath -Force -Recurse -Confirm:$false -ErrorAction Stop
                }
                Write-Info -Message "Removed: $registryPath" -Color ([ConsoleColor]::Green)
            } catch {
                Write-Error "Failed to remove: $registryPath - $($_.Exception.Message)"
            }
        }
        Write-Info -Message "Deep-clean completed." -Color ([ConsoleColor]::Green)
    } else {
        Write-Info -Message "Deep-clean helpers not available, skipping..." -Color ([ConsoleColor]::Yellow)
    }
}

function Download-And-Extract {
    param([string]$EffectiveBranch)

    $zipUrl = "https://github.com/Joly0/Run-in-Sandbox/archive/refs/heads/$EffectiveBranch.zip"
    $tempPath = [IO.Path]::GetTempPath()
    $zipPath = Join-Path $tempPath "Run-in-Sandbox-$EffectiveBranch.zip"
    $extractPath = Join-Path $tempPath "Run-in-Sandbox-$EffectiveBranch"

    if (Test-Path $extractPath) {
        Write-Verbose "Removing existing extracted folder..."
        Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    try {
        Write-Verbose ("Downloading from branch '{0}'..." -f $EffectiveBranch)
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing -TimeoutSec 120
        $ProgressPreference = 'Continue'
        Write-Verbose "Download completed: $zipPath"
    } catch {
        throw "Download failed: $($_.Exception.Message)"
    }

    try {
        Write-Verbose "Extracting package..."
        $ProgressPreference = 'SilentlyContinue'
        Expand-Archive -Path $zipPath -DestinationPath $tempPath -Force
        $ProgressPreference = 'Continue'
        Write-Verbose "Extraction completed to $tempPath"
    } catch {
        throw "Extraction failed: $($_.Exception.Message)"
    }

    try { Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue } catch {}

    # The extracted structure is: Run-in-Sandbox-branch/Sources/Run_in_Sandbox/
    # We need to return the path to the extracted folder (which contains the Sources folder)
    return $extractPath
}

function Get-DefaultStartupScriptNames {
    param([string]$ExtractPath)
    # The extracted structure is: ExtractPath/Sources/Run_in_Sandbox/startup-scripts
    $defaultStartupDir = Join-Path $ExtractPath "Sources\Run_in_Sandbox\startup-scripts"
    if (Test-Path $defaultStartupDir) {
        return (Get-ChildItem -Path $defaultStartupDir -File | Select-Object -ExpandProperty Name)
    }
    return @()
}

function Backup-CustomStartupScripts {
    param(
        [string]$RunFolder,
        [string[]]$DefaultNames
    )

    $installedStartupDir = Join-Path $RunFolder "\startup-scripts"
    if (Test-Path $installedStartupDir) {
        Get-ChildItem -Path $installedStartupDir -File |
            Where-Object { $DefaultNames -notcontains $_.Name } |
            ForEach-Object {
                Copy-Item $_.FullName (Join-Path $script:customScriptsBackupDir $_.Name) -Force
            }
    }
}

function Remove-OldInstallIfDeepClean {
    param(
        [switch]$DeepClean,
        [string]$RunFolder,
        [string[]]$DefaultNames
    )
    if (-not $DeepClean) { return }

    try { Backup-CustomStartupScripts -RunFolder $RunFolder -DefaultNames $DefaultNames } catch {}
    if (Test-Path $RunFolder) {
        try { Remove-Item -Path $RunFolder -Recurse -Force } catch {}
    }
    New-Item -ItemType Directory -Path $RunFolder -Force | Out-Null
}

function Restore-CustomStartupScripts {
    param([string]$RunFolder)

    if (-not $script:customScriptsBackupDir) { return }
    $installedStartupDir = Join-Path $RunFolder "\startup-scripts"
    if (-not (Test-Path $installedStartupDir)) {
        New-Item -ItemType Directory -Path $installedStartupDir -Force | Out-Null
    }

    if (Test-Path $script:customScriptsBackupDir) {
        Get-ChildItem -Path $script:customScriptsBackupDir -File | ForEach-Object {
            Copy-Item $_.FullName (Join-Path $installedStartupDir $_.Name) -Force
        }
        try { Remove-Item $script:customScriptsBackupDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Sync-CoreFiles {
    param(
        [string]$ExtractPath,
        [string]$RunFolder,
        [string[]]$DefaultNames
    )
    # Copy core root-level files if present in the extracted package
    # These files are in the root of the extracted folder (e.g., Run-in-Sandbox-master/)
    $rootFiles = @(
        "CommonFunctions.ps1",
        "version.json"
    )
    
    foreach ($fileName in $rootFiles) {
        $sourceFile = Join-Path $ExtractPath $fileName
        if (Test-Path $sourceFile) {
            Copy-Item $sourceFile (Join-Path $RunFolder $fileName) -Force
            Write-Verbose "Copied $fileName from root of extracted package"
        }
    }
    
    # Copy Sources except startup-scripts (handled separately below)
    # The extracted structure is: ExtractPath/Sources/Run_in_Sandbox/
    $sourceSandboxDir = Join-Path $ExtractPath "Sources\Run_in_Sandbox"
    # After installation, files are directly in $RunFolder, not in a subfolder
    $destSandboxDir   = $RunFolder
    
    Write-Verbose "Looking for Sources in: $sourceSandboxDir"
    
    if (Test-Path $sourceSandboxDir) {
        Write-Verbose "Found Sources folder, syncing contents..."
        Get-ChildItem -Path $sourceSandboxDir | Where-Object { $_.Name -notin @("startup-scripts", "Sandbox_Config.xml") } | ForEach-Object {
            $destPath = Join-Path $destSandboxDir $_.Name
            if ($_.PSIsContainer) {
                # If destination directory exists, sync contents instead of copying the whole directory
                if (Test-Path $destPath) {
                    # Sync the contents of the directory
                    Get-ChildItem -Path $_.FullName -Recurse | ForEach-Object {
                        $relativePath = $_.FullName.Substring($sourceSandboxDir.Length)
                        $finalDestPath = Join-Path $destSandboxDir $relativePath
                        # Ensure parent directory exists
                        $parentDir = Split-Path $finalDestPath -Parent
                        if (-not (Test-Path $parentDir)) {
                            New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                        }
                        Copy-Item $_.FullName $finalDestPath -Force
                    }
                } else {
                    # If destination doesn't exist, create it and copy contents
                    Copy-Item $_.FullName $destPath -Recurse -Force
                }
            } else {
                Copy-Item $_.FullName $destPath -Force
            }
        }
        
        # Handle Sandbox_Config.xml - only copy if it doesn't exist in destination (preserve user changes)
        $sourceConfig = Join-Path $sourceSandboxDir "Sandbox_Config.xml"
        $destConfig = Join-Path $destSandboxDir "Sandbox_Config.xml"
        if ((Test-Path $sourceConfig) -and (-not (Test-Path $destConfig))) {
            Copy-Item $sourceConfig $destConfig -Force
        }
        
        # Update startup scripts - merge default scripts while preserving user-added ones
        $defaultStartupDir = Join-Path $sourceSandboxDir "startup-scripts"
        if (Test-Path $defaultStartupDir) {
            $destStartupDir = Join-Path $destSandboxDir "startup-scripts"
            if (-not (Test-Path $destStartupDir)) {
                New-Item -ItemType Directory -Path $destStartupDir -Force | Out-Null
            }
            
            # Copy all default scripts, but don't overwrite existing ones with the same name
            Get-ChildItem -Path $defaultStartupDir -File | ForEach-Object {
                $destScript = Join-Path $destStartupDir $_.Name
                if (-not (Test-Path $destScript)) {
                    Copy-Item $_.FullName $destScript -Force
                }
            }
        }
    } else {
        Write-Verbose "Sources folder not found at expected location: $sourceSandboxDir"
    }
}

function Run-AddStructure {
    param(
        [string]$ExtractPath,
        [switch]$NoCheckpoint,
        [bool]$IsInstalled
    )

    # Unblock and set execution policy for session
    try {
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force -ErrorAction SilentlyContinue
    } catch { }
    Get-ChildItem -LiteralPath $ExtractPath -Recurse -Filter "*.ps1" | Unblock-File

    $scriptPath = Join-Path $ExtractPath "Add_Structure.ps1"
    if (-not (Test-Path $scriptPath)) {
        throw "Add_Structure.ps1 not found in extracted content."
    }

    if ($IsInstalled) {
        Write-Info "Updating Run-in-Sandbox..." ([ConsoleColor]::Cyan)
    } else {
        Write-Info "Installing Run-in-Sandbox..." ([ConsoleColor]::Cyan)
    }

    # For new installations, we need to copy the Sources folder to ProgramData first
    # The Add_Structure.ps1 script expects Sources to be in the same directory as the script
    $sourcesPath = Join-Path $ExtractPath "Sources"
    if (-not $IsInstalled -and (Test-Path $sourcesPath)) {
        Write-Verbose "Copying Sources folder to ProgramData for initial installation..."
        try {
            # Ensure the Run_in_Sandbox folder exists
            if (-not (Test-Path $Run_in_Sandbox_Folder)) {
                New-Item -Path $Run_in_Sandbox_Folder -ItemType Directory -Force | Out-Null
            }
            
            # Copy the Sources folder to ProgramData
            Copy-Item -Path $sourcesPath -Destination $env:ProgramData -Force -Recurse
            Write-Verbose "Sources folder copied to ProgramData"
        } catch {
            Write-Verbose "Failed to copy Sources folder: $($_.Exception.Message)"
        }
    }

    Push-Location $ExtractPath
    try {
        if ($NoCheckpoint) {
            & ".\Add_Structure.ps1" -NoCheckpoint
        } else {
            & ".\Add_Structure.ps1"
        }
    } finally {
        Pop-Location
    }
}

function Cleanup-Temp {
    param([string]$ExtractPath, [string]$RunFolder)

    if (Test-Path $ExtractPath) {
        Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    $StateFile = Join-Path $RunFolder "temp\UpdateState.json"
    if (Test-Path $StateFile) { Remove-Item $StateFile -Force -ErrorAction SilentlyContinue }
}

Export-ModuleMember -Function @(
    'Request-OptionalDeepClean',
    'New-InstallBackup',
    'Invoke-DeepCleanIfRequested',
    'Download-And-Extract',
    'Get-DefaultStartupScriptNames',
    'Backup-CustomStartupScripts',
    'Remove-OldInstallIfDeepClean',
    'Restore-CustomStartupScripts',
    'Sync-CoreFiles',
    'Run-AddStructure',
    'Cleanup-Temp'
)