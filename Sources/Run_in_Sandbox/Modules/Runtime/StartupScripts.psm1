<#
.SYNOPSIS
    Startup Scripts module for Run-in-Sandbox

.DESCRIPTION
    This module provides startup script management functionality for the Run-in-Sandbox application.
    It handles execution and management of scripts that run when the sandbox starts.
#>

function Enable-StartupScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OriginalCommand,  # whatever you would have passed as -Command_to_Run
        [string]$StartupScriptFolderName = "startup-scripts"
    )

    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $Sandbox_Root_Path = "C:\Run_in_Sandbox"
    
    $StartupScriptsFolder = Join-Path $Run_in_Sandbox_Folder $StartupScriptFolderName
    New-Item -ItemType Directory -Path $StartupScriptsFolder -Force | Out-Null
    
    if ($OriginalCommand -ne "") {
       # Write the original command into a file
        $origCmdFile = Join-Path $StartupScriptsFolder "OriginalCommand.txt"
        Set-Content -LiteralPath $origCmdFile -Value $OriginalCommand -Encoding UTF8 -Force 
    }

    # Orchestrator that runs NN-*.ps1 in lexicographic order, then runs the original command
    $orchestrator = @'
param(
    [string]$ScriptsPath = "C:\Run_in_Sandbox\startup-scripts",
    # Can include this switch when running from the .wsb file to indicate it's the first launch of the sandbox
    # Useful if re-running this script within the sandbox as a test, but don't want certain parts to run again
    [switch]$launchingSandbox
)

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-host "`n`nERROR: This script is intended to be run from WITHIN the Windows Sandbox.`nIt appears you are running this from outside the sandbox.`n" -ForegroundColor Red
    Write-host "`nPress Enter to exit." -ForegroundColor Yellow
    Read-Host
    exit
}

Write-Host "[Orchestrator] Scripts path: $ScriptsPath"

# 1) Run ordered startup scripts: 00-*, 01-* ... 99-*
$pattern = '^\d{2}-.+\.ps1$'
$items = Get-ChildItem -LiteralPath $ScriptsPath -Filter *.ps1 -File -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -match $pattern } |
         Sort-Object Name

foreach ($i in $items) {
    Write-Host "[Orchestrator] Running: $($i.Name)"
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $i.FullName
        $rc = $LASTEXITCODE
        if ($rc -ne $null -and $rc -ne 0) {
            Write-Warning "[Orchestrator] Script $($i.Name) returned exit code $rc"
        }
    } catch {
        Write-Warning "[Orchestrator] Script $($i.Name) threw: $($_.Exception.Message)"
    }
}

# Restart Explorer so changes take effect
Write-Host "[Orchestrator] Restarting Explorer so changes take effect"
Get-Process explorer | Stop-Process -Force

# 2) Read and run the original command last
$origFile = Join-Path $ScriptsPath "OriginalCommand.txt"
if (Test-Path -LiteralPath $origFile) {
    $orig = Get-Content -LiteralPath $origFile -Raw
    Write-Host "[Orchestrator] Running original command..."
    # Run through cmd to support both cmd and PowerShell-style lines
    Start-process -Filepath "C:\Windows\SysWOW64\cmd.exe" -ArgumentList @('/c', '"' + $orig + '"') -WindowStyle Hidden
} else {
    Write-Warning "[Orchestrator] OriginalCommand.txt not found; nothing to run."
}
'@

    $orchestratorPath = Join-Path $StartupScriptsFolder "_orchestrator.ps1"
    Set-Content -LiteralPath $orchestratorPath -Value $orchestrator -Encoding UTF8 -Force

    # Return the single Sandbox command that runs the orchestrator
    "C:\Run_in_Sandbox\ServiceUI.exe -Process:explorer.exe C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Normal -NoProfile -ExecutionPolicy Bypass -NoExit -File `"$Sandbox_Root_Path\$StartupScriptFolderName\_orchestrator.ps1`""
}


function Add-NotepadToSandbox {
    [CmdletBinding()]
    param(
        [string]$HostPayloadRoot = "C:\ProgramData\Run_in_Sandbox\NotepadPayload",
        [switch]$EnforceEnUsFallback # if set, will try en-US when preferred language is missing
    )

    # Resolve a single notepad.exe (prefer System32)
    $exeCandidates = Get-Command notepad.exe -ErrorAction Stop | Select-Object -ExpandProperty Source
    $exePath = ($exeCandidates | Where-Object { $_ -match '\\Windows\\System32\\' } | Select-Object -First 1)
    if (-not $exePath) { $exePath = $exeCandidates | Select-Object -First 1 }

    $exeDir  = Split-Path $exePath -Parent
    $exeName = Split-Path $exePath -Leaf

    # Build candidate language list
    $candidates = @()
    try { $candidates += (Get-UICulture).Name } catch {}
    try { $candidates += (Get-WinSystemLocale).Name } catch {}
    try {
        $candidates += Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\MUI\UILanguages" |
                       Select-Object -ExpandProperty PSChildName
    } catch {}
    $candidates = $candidates | Where-Object { $_ } | Select-Object -Unique

    # Probe possible MUI locations
    $dirs = @(
        $exeDir,
        (Join-Path $env:WINDIR 'System32'),
        (Join-Path $env:WINDIR 'SysWOW64'),
        $env:WINDIR
    ) | Select-Object -Unique

    $muiPath = $null; $resolvedLang = $null
    foreach ($lang in $candidates) {
        foreach ($dir in $dirs) {
            $p = Join-Path (Join-Path $dir $lang) "$exeName.mui"
            if (Test-Path -LiteralPath $p) { $muiPath = $p; $resolvedLang = $lang; break }
        }
        if ($muiPath) { break }
    }

    if (-not $muiPath -and $EnforceEnUsFallback) {
        foreach ($dir in $dirs) {
            $fallback = Join-Path (Join-Path $dir 'en-US') "$exeName.mui"
            if (Test-Path -LiteralPath $fallback) { $muiPath = $fallback; $resolvedLang = 'en-US'; break }
        }
    }

    if (-not $muiPath) {
        throw "Could not locate notepad.exe.mui for $exePath. On some systems Notepad is a Store app without a classic MUI."
    }

    # Stage payload on host: System32\notepad.exe and System32\<lang>\notepad.exe.mui
    $sys32Out = Join-Path $HostPayloadRoot "System32"
    $langOut  = Join-Path $sys32Out $resolvedLang
    New-Item -ItemType Directory -Path $langOut -Force | Out-Null
    Copy-Item -LiteralPath $exePath -Destination (Join-Path $sys32Out $exeName) -Force
    Copy-Item -LiteralPath $muiPath -Destination (Join-Path $langOut "$exeName.mui") -Force
}

Export-ModuleMember -Function @(
    'Enable-StartupScripts',
    'Add-NotepadToSandbox'
)
