param (
    [String]$Intunewin_Content_File = "C:\Run_in_Sandbox\temp\Intunewin_Folder.txt",
    [String]$Intunewin_Command_File = "C:\Run_in_Sandbox\temp\Intunewin_Install_Command.txt"
)
if (-not (Test-Path $Intunewin_Content_File) ) {
	EXIT
}
if (-not (Test-Path $Intunewin_Command_File) ) {
	EXIT
}

$Sandbox_Folder = "C:\Run_in_Sandbox"

# Read Hide_Powershell configuration from Sandbox_Config.xml
$ConfigPath = "$Sandbox_Folder\Sandbox_Config.xml"
$Hide_Powershell = "True"  # Default value
if (Test-Path $ConfigPath) {
    $config = [xml](Get-Content $ConfigPath)
    $Hide_Powershell = $config.Configuration.Hide_Powershell
    if ([string]::IsNullOrEmpty($Hide_Powershell)) { 
        $Hide_Powershell = "True" 
    }
}

# Always use /c for non-blocking execution
# Error handling will show issues to user when Hide_Powershell is False
$CmdSwitch = "/c"
$ShowErrors = ($Hide_Powershell -eq "False")
$ScriptPath = Get-Content -Raw $Intunewin_Content_File
$Command = Get-Content -Raw $Intunewin_Command_File
$Command = $Command.replace('"','')

$FileName = (Get-Item $ScriptPath).BaseName

$Intunewin_Extracted_Folder = "C:\Windows\Temp\intunewin"
New-Item -Path $Intunewin_Extracted_Folder -Type Directory -Force | Out-Null
Copy-Item -Path $ScriptPath -Destination $Intunewin_Extracted_Folder -Force
$New_Intunewin_Path = "$Intunewin_Extracted_Folder\$FileName.intunewin"

Set-Location -Path $Sandbox_Folder
& .\IntuneWinAppUtilDecoder.exe $New_Intunewin_Path -s
$IntuneWinDecoded_File_Name = "$Intunewin_Extracted_Folder\$FileName.decoded.zip"

New-Item -Path "$Intunewin_Extracted_Folder\$FileName" -Type Directory -Force | Out-Null

$IntuneWin_Rename = "$FileName.zip"

Rename-Item -Path $IntuneWinDecoded_File_Name -NewName $IntuneWin_Rename -Force

$Extract_Path = "$Intunewin_Extracted_Folder\$FileName"
Expand-Archive -LiteralPath "$Intunewin_Extracted_Folder\$IntuneWin_Rename" -DestinationPath $Extract_Path -Force

Remove-Item -Path "$Intunewin_Extracted_Folder\$IntuneWin_Rename" -Force
Start-Sleep -Seconds 1

$ServiceUI = "$Sandbox_Folder\ServiceUI.exe"
$PsExec = "$Sandbox_Folder\PsExec.exe"

# Load Windows Forms for error dialogs
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# Create a hidden topmost form to use as MessageBox owner (ensures dialogs appear on top)
$topmostForm = New-Object System.Windows.Forms.Form
$topmostForm.TopMost = $true
$topmostForm.ShowInTaskbar = $false
$topmostForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized

# Create a batch file with debug info and the actual command
# This avoids issues with & characters in command line arguments
$BatchFile = "$Extract_Path\_install.cmd"
$BatchContent = @"
@echo off
echo ==========================================
echo [SOURCE: IntuneWin_Install.ps1]
echo Installing: $FileName
echo Command: $Command
echo Working Dir: $Extract_Path
echo ==========================================
$Command
exit /b %errorlevel%
"@
Set-Content -LiteralPath $BatchFile -Value $BatchContent -Encoding ASCII -Force

# Execute the batch file - PsExec handles the command line properly
# The CMD window is visible when Hide_Powershell is False (via ServiceUI)
& $PsExec \\localhost -w "$Extract_Path" -nobanner -accepteula -s $ServiceUI -Process:explorer.exe C:\Windows\SysWOW64\cmd.exe $CmdSwitch $BatchFile
$exitCode = $LASTEXITCODE

# Clean up batch file
Remove-Item -LiteralPath $BatchFile -Force -ErrorAction SilentlyContinue

# Show error dialog if execution failed and Hide_Powershell is False
if ($exitCode -ne 0 -and $ShowErrors) {
    $topmostForm.Show()
    $topmostForm.Activate()
    [System.Windows.Forms.MessageBox]::Show(
        $topmostForm,
        "IntuneWin installation failed: $FileName`nExit code: $exitCode`nCommand: $Command",
        "Installation Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    $topmostForm.Hide()
}

# Open debug console when Hide_Powershell is False (for debugging purposes)
# Use /k to keep the window open for inspection
if ($ShowErrors) {
    $DebugBatchFile = "$Extract_Path\_debug.cmd"
    $DebugContent = @"
@echo off
echo ==========================================
echo [DEBUG CONSOLE: IntuneWin_Install.ps1]
echo Package: $FileName
echo Working Dir: %CD%
echo ==========================================
echo.
echo Type 'dir' to list files, 'exit' to close.
"@
    Set-Content -LiteralPath $DebugBatchFile -Value $DebugContent -Encoding ASCII -Force
    & $PsExec \\localhost -w "$Extract_Path" -nobanner -accepteula -s -d $ServiceUI -Process:explorer.exe C:\Windows\SysWOW64\cmd.exe /k $DebugBatchFile
}

# Return exit code so parent script can handle it
exit $exitCode