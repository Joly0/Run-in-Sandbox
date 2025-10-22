#Requires -Version 5.1
[CmdletBinding()]
param (
    [switch]$NoCheckpoint,
    [switch]$DeepClean,
    [switch]$AutoUpdate,
    [string]$Branch = ""
)

# ======================================================================================
# Minimal, maintainable installer with functions and optional verbose logging
# - Keep functionality intact
# - Reduce console noise (use Write-Verbose for debug)
# - Simplify elevation and download flow
# ======================================================================================

# Globals
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$IsInstalled = Test-Path $Run_in_Sandbox_Folder

function Write-Info {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    # Only show to end-user for relevant messages (not too chatty)
    Write-Host $Message -ForegroundColor $Color
}

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

function Test-IsAdmin {
    $windowsIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $windowsPrincipal = [Security.Principal.WindowsPrincipal]::new($windowsIdentity)
    return $windowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-Admin {
    param(
        [string]$EffectiveBranch,
        [switch]$NoCheckpoint,
        [switch]$DeepClean,
        [switch]$AutoUpdate
    )
    if (Test-IsAdmin) {
        Write-Verbose "Already elevated."
        return
    }

    Write-Info "Elevation required. Restarting installer as Administrator..." ([ConsoleColor]::Yellow)

    # Simplified, reliable elevation: download current branch installer to temp and re-run with same parameters.
    $TempScript = Join-Path ([IO.Path]::GetTempPath()) "Install_Run-in-Sandbox.Elevated.ps1"
    $InstallerUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$EffectiveBranch/Install_Run-in-Sandbox.ps1"

    try {
        Write-Verbose "Downloading elevated installer from: $InstallerUrl"
        Invoke-WebRequest -Uri $InstallerUrl -UseBasicParsing -OutFile $TempScript -TimeoutSec 45 -ErrorAction Stop
    } catch {
        Write-Info "Failed to download installer for elevation: $($_.Exception.Message)" ([ConsoleColor]::Red)
        Read-Host "Press Enter to exit"
        break script
    }

    $argsList = @(
        "-NoExit",
        "-ExecutionPolicy", "Bypass",
        "-File", $TempScript
    )
    if ($NoCheckpoint) { $argsList += "-NoCheckpoint" }
    if ($DeepClean)    { $argsList += "-DeepClean" }
    if ($AutoUpdate)   { $argsList += "-AutoUpdate" }
    if ($EffectiveBranch) {
        $argsList += @("-Branch", $EffectiveBranch)
    }

    Write-Verbose ("Elevation command: powershell.exe " + ($argsList -join " "))

    try {
        Start-Process powershell.exe -ArgumentList $argsList -Verb RunAs | Out-Null
    } catch {
        Write-Info "Elevation failed: $($_.Exception.Message)" ([ConsoleColor]::Red)
        Read-Host "Press Enter to exit"
        break script
    }

    break script
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

function Prompt-OptionalDeepClean {
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

    return $extractPath
}
function Get-DefaultStartupScriptNames {
    param([string]$ExtractPath)
    $defaultStartupDir = Join-Path $ExtractPath "\startup-scripts"
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
    $script:customScriptsBackupDir = Join-Path $env:TEMP "RIS_CustomStartupScripts"
    try {
        if (Test-Path $script:customScriptsBackupDir) {
            Remove-Item $script:customScriptsBackupDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {}
    New-Item -ItemType Directory -Force -Path $script:customScriptsBackupDir | Out-Null

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
    $rootFiles = @(
        "CommonFunctions.ps1",
        "version.json"
    )
    foreach ($fileName in $rootFiles) {
        $sourceFile = Join-Path $ExtractPath $fileName
        if (Test-Path $sourceFile) {
            Copy-Item $sourceFile (Join-Path $RunFolder $fileName) -Force
        }
    }

    # Copy Sources except startup-scripts (handled separately below)
    $sourceSandboxDir = Join-Path $ExtractPath "Sources\Run_in_Sandbox"
    # After installation, files are directly in $RunFolder, not in a subfolder
    $destSandboxDir   = $RunFolder
    if (Test-Path $sourceSandboxDir) {
        Get-ChildItem -Path $sourceSandboxDir | Where-Object { $_.Name -notin @("startup-scripts", "Sandbox_Config.xml") } | ForEach-Object {
            $destPath = Join-Path $destSandboxDir $_.Name
            if ($_.PSIsContainer) {
                Copy-Item $_.FullName $destPath -Recurse -Force
            } else {
                Copy-Item $_.FullName $destPath -Force
            }
        }

        # Handle Sandbox_Config.xml - only copy if it doesn't exist in destination (preserve user changes)
        $sourceConfig = Join-Path $sourceSandboxDir "Sandbox_Config.xml"
        $destConfig = Join-Path $destSandboxDir "Sandbox_Config.xml"
        if (Test-Path $sourceConfig -and -not (Test-Path $destConfig)) {
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

function Merge-ConfigIfNeeded {
    param(
        [bool]$IsInstalled,
        [string]$RunFolder
    )

    if (-not $IsInstalled) { return }

    $ConfigBackup = "$env:TEMP\Sandbox_Config_Backup.xml"
    if (Test-Path "$RunFolder\Sandbox_Config.xml") {
        Copy-Item "$RunFolder\Sandbox_Config.xml" $ConfigBackup -Force
    } else {
        return
    }

    try {
        Write-Info "Merging configuration..." ([ConsoleColor]::Cyan)
        Merge-SandboxConfig -OldConfigPath $ConfigBackup -NewConfigPath "$RunFolder\Sandbox_Config.xml"
    } catch {
        Write-Verbose "Merge-SandboxConfig failed, attempting simple merge: $($_.Exception.Message)"
        try {
            [xml]$oldConfig = Get-Content $ConfigBackup
            [xml]$newConfig = Get-Content "$RunFolder\Sandbox_Config.xml"
            if ($oldConfig.Configuration -and $newConfig.Configuration) {
                foreach ($child in $oldConfig.Configuration.ChildNodes) {
                    if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                    $target = $null
                    foreach ($n in $newConfig.Configuration.ChildNodes) {
                        if ($n.NodeType -eq [System.Xml.XmlNodeType]::Element -and $n.Name -eq $child.Name) {
                            $target = $n
                            break
                        }
                    }
                    if ($target) {
                        $target.InnerText = $child.InnerText
                    }
                }
                $newConfig.Save("$RunFolder\Sandbox_Config.xml") | Out-Null
            }
        } catch {
            Write-Verbose "Simple merge failed: $($_.Exception.Message)"
        }
    } finally {
        try { Remove-Item $ConfigBackup -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Validate-Installation {
    param([string]$RunFolder)

    $RequiredFiles = @(
        (Join-Path $RunFolder "RunInSandbox.ps1"),
        (Join-Path $RunFolder "CommonFunctions.ps1"),
        (Join-Path $RunFolder "Sandbox_Config.xml"),
        (Join-Path $RunFolder "version.json")
    )

    $ok = $true
    foreach ($f in $RequiredFiles) {
        if (-not (Test-Path $f)) {
            Write-Info "Validation failed: Missing $f" ([ConsoleColor]::Red)
            $ok = $false
        }
    }
    return $ok
}

function Cleanup-Temp {
    param([string]$ExtractPath, [string]$RunFolder)

    if (Test-Path $ExtractPath) {
        Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    $StateFile = Join-Path $RunFolder "temp\UpdateState.json"
    if (Test-Path $StateFile) { Remove-Item $StateFile -Force -ErrorAction SilentlyContinue }
}

# -------------------------------------------------------------------------------------------------
# Branch resolution and elevation
# -------------------------------------------------------------------------------------------------
$Branch = Resolve-Branch -Requested $Branch -Installed:$IsInstalled
Write-Verbose "Effective branch: $Branch"

Ensure-Admin -EffectiveBranch $Branch -NoCheckpoint:$NoCheckpoint -DeepClean:$DeepClean -AutoUpdate:$AutoUpdate

# -------------------------------------------------------------------------------------------------
# Load common functions (prefer online for current branch, fallback to local) in script scope
# -------------------------------------------------------------------------------------------------
try {
    $commonFunctionsUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$Branch/CommonFunctions.ps1"
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
    $LatestVersion  = if (-not $AutoUpdate) { Get-LatestVersionFromBranch -EffectiveBranch $Branch } else { $null }
    $InstalledBranch = Get-InstalledBranch

    if (-not $AutoUpdate -and $CurrentVersion) {
        Write-Info ("Current Version:  {0}" -f $CurrentVersion) ([ConsoleColor]::Green)
        $branchToShow = if ($InstalledBranch) { $InstalledBranch } else { $Branch }
        Write-Info ("Current Branch:   {0}" -f $branchToShow) ([ConsoleColor]::Cyan)
        if ($LatestVersion) {
            Write-Info ("Latest Version:   {0}" -f $LatestVersion) ([ConsoleColor]::Green)
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
    $DeepClean = Prompt-OptionalDeepClean -AutoUpdate:$AutoUpdate -DeepCleanRef:$DeepClean
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
    $extractPath = Download-And-Extract -EffectiveBranch $Branch
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
    Run-AddStructure -ExtractPath $extractPath -NoCheckpoint:$NoCheckpoint -IsInstalled:$IsInstalled
    Write-Host "Running Merge-ConfigIfNeeded with parameters $IsInstalled and $Run_in_Sandbox_Folder"
    Merge-ConfigIfNeeded -IsInstalled:$IsInstalled -RunFolder:$Run_in_Sandbox_Folder
    Write-Host "Running Sync-CoreFiles with parameters $Run_in_Sandbox_Folder and $DefaultStartupNames"
    Sync-CoreFiles -ExtractPath $extractPath -RunFolder $Run_in_Sandbox_Folder -DefaultNames $DefaultStartupNames
    Write-Host "Running Restore-CustomStartupScripts with parameters $Run_in_Sandbox_Folder"
    Restore-CustomStartupScripts -RunFolder $Run_in_Sandbox_Folder
    Write-Host "Running Ensure-VersionJson with parameters $Run_in_Sandbox_Folder and $extractPath and $Branch and $LatestVersion"
    Ensure-VersionJson -RunFolder $Run_in_Sandbox_Folder -ExtractPath $extractPath -EffectiveBranch $Branch -LatestVersion $LatestVersion

    $valid = Validate-Installation -RunFolder $Run_in_Sandbox_Folder
    if (-not $valid) {
        Write-Info "Installation validation failed!" ([ConsoleColor]::Red)
        if ($BackupCreated) {
            Write-Info "Backup available at: $Run_in_Sandbox_Folder\backup" ([ConsoleColor]::Yellow)
        }
        break script
    }

    Cleanup-Temp -ExtractPath $extractPath -RunFolder $Run_in_Sandbox_Folder

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
    Read-Host "Press Enter to exit"
}
