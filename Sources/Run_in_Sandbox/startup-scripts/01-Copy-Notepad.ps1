# Runs inside Windows Sandbox. Assumes the host mapped ProgramData\Run_in_Sandbox to C:\Run_in_Sandbox in the Sandbox [1].

$srcRoot = "C:\Run_in_Sandbox\NotepadPayload"
$srcSys  = Join-Path $srcRoot "System32"
$dstSys  = "C:\Windows\System32"
$dst  = "C:\Windows"

$hadError = $false
$notepadStaged = $false

# 1) Copy notepad.exe if present
try {
    $srcExe = Join-Path $srcSys "notepad.exe"
    $notepadPathSys = Join-Path $dstSys "notepad.exe"
    $notepadPath = Join-Path $dst "notepad.exe"
    if (Test-Path -LiteralPath $srcExe) {
        Copy-Item -LiteralPath $srcExe -Destination $notepadPath -Force
        Copy-Item -LiteralPath $srcExe -Destination $notepadPathSys -Force
        $notepadStaged = $true
        Write-Host "[Copy-Notepad] notepad.exe copied."
    } else {
        Write-Warning "[Copy-Notepad] Source not found: $srcExe"
    }
} catch {
    Write-Warning "[Copy-Notepad] Failed to copy notepad.exe: $($_.Exception.Message)"
    $hadError = $true
}

# 2) Copy the MUI file from the language subfolder(s) (e.g., en-US, de-DE)
try {
    $langDirs = Get-ChildItem -LiteralPath $srcSys -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^[a-z]{2}-[A-Z]{2}$' }

    foreach ($langDir in $langDirs) {
        $srcMui = Join-Path $langDir.FullName "notepad.exe.mui"
        if (Test-Path -LiteralPath $srcMui) {
            $dstLangDir = Join-Path $dst $langDir.Name
            $dstLangDirSys = Join-Path $dstSys $langDir.Name
            if (-not (Test-Path -LiteralPath $dstLangDir)) {
                New-Item -ItemType Directory -Path $dstLangDir -Force | Out-Null
            }
            if (-not (Test-Path -LiteralPath $dstLangDirSys)) {
                New-Item -ItemType Directory -Path $dstLangDirSys -Force | Out-Null
            }
            $dstMui = Join-Path $dstLangDir "notepad.exe.mui"
            $dstMuiSys = Join-Path $dstLangDirSys "notepad.exe.mui"
            if (-not (Test-Path -LiteralPath $dstMui) ) {
                Copy-Item -LiteralPath $srcMui -Destination $dstMui -Force
                Write-Host "[Copy-Notepad] MUI copied for $($langDir.Name)."
            }
            if (-not (Test-Path -LiteralPath $dstMuiSys) ) {
                Copy-Item -LiteralPath $srcMui -Destination $dstMuiSys -Force
                Write-Host "[Copy-Notepad] MUI copied for $($langDir.Name)."
            }
        } else {
            Write-Warning "[Copy-Notepad] MUI not found in: $($langDir.FullName)"
        }
    }
} catch {
    Write-Warning "[Copy-Notepad] Failed to copy MUI: $($_.Exception.Message)"
    $hadError = $true
}

if ($hadError) {
    exit 1
} elseif (-not $notepadStaged) {
    exit 0
} else {
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad" /f
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad" /v "Icon" /t REG_SZ /d "$notepadPath,0" /f
	reg add "HKEY_CLASSES_ROOT\*\shell\Edit with Notepad\command" /ve /d "`"$notepadPath`" `"%1`"" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad" /ve /d "Open Notepad" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad" /v "Icon" /t REG_SZ /d "$notepadPath,0" /f
	reg add "HKEY_CLASSES_ROOT\Directory\Background\shell\Notepad\command" /ve /d "`"$notepadPath`"" /f
    cmd /c assoc .txt=txtfile
    If ( -not (Test-Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command')) { New-Item -Path 'HKLM:\SOFTWARE\Classes\txtfile\shell\open\command' -Force }
    cmd /c ftype txtfile=`"$notepadPath`" "%1"
    
    # Restart Explorer so changes take effect
    # Get-Process explorer | Stop-Process -Force
    # Open an explorer window to the host-shared folder on first launch
    # Start-Process explorer.exe C:\Users\WDAGUtilityAccount\Desktop\HostShared
    
    exit 0
}