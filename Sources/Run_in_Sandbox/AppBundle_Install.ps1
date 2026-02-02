$Desktop = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$App_Bundle_File = "$Sandbox_Root_Path\App_Bundle.sdbapp"
$LogFolder = "$Desktop\AppBundle_Logs"

# Read Hide_Powershell configuration from Sandbox_Config.xml
$ConfigPath = "$Sandbox_Root_Path\Sandbox_Config.xml"
$Hide_Powershell = "True"  # Default value
if (Test-Path $ConfigPath) {
    $config = [xml](Get-Content $ConfigPath)
    $Hide_Powershell = $config.Configuration.Hide_Powershell
    if ([string]::IsNullOrEmpty($Hide_Powershell)) { 
        $Hide_Powershell = "True" 
    }
}

# Calculate WindowStyle based on configuration
$WindowStyle = if ($Hide_Powershell -eq "False") { "Normal" } else { "Hidden" }
$ShowErrors = ($Hide_Powershell -eq "False")

# Show debug header when Hide_Powershell is False
if ($ShowErrors) {
    # Set window title and increase buffer size to prevent scrolling issues
    $Host.UI.RawUI.WindowTitle = "AppBundle Installer"
    try {
        $bufferSize = $Host.UI.RawUI.BufferSize
        $bufferSize.Height = 3000
        $Host.UI.RawUI.BufferSize = $bufferSize
    } catch {
        # Ignore buffer size errors
    }
    
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "[SOURCE: AppBundle_Install.ps1]" -ForegroundColor Cyan
    Write-Host "App Bundle File: $App_Bundle_File" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
}

# Load Windows Forms for error dialogs
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# Create a hidden topmost form to use as MessageBox owner (ensures dialogs appear on top)
$topmostForm = New-Object System.Windows.Forms.Form
$topmostForm.TopMost = $true
$topmostForm.ShowInTaskbar = $false
$topmostForm.WindowState = [System.Windows.Forms.FormWindowState]::Minimized

# Helper function to show error when Hide_Powershell is False
function Show-InstallationError {
    param(
        [string]$AppName,
        [string]$FilePath,
        [int]$ExitCode,
        [string]$ErrorMessage = ""
    )
    
    if (-not $ShowErrors) { return }
    
    $message = "Installation failed: $AppName`n"
    $message += "File: $FilePath`n"
    $message += "Exit code: $ExitCode"
    if ($ErrorMessage) {
        $message += "`nError: $ErrorMessage"
    }
    
    # Show the form briefly to ensure MessageBox appears on top
    $topmostForm.Show()
    $topmostForm.Activate()
    
    [System.Windows.Forms.MessageBox]::Show(
        $topmostForm,
        $message,
        "Installation Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    
    $topmostForm.Hide()
}

$SDBApp_Root_Path = "C:\SBDApp"
$Get_Apps_to_install = [xml](Get-Content $App_Bundle_File)
$Apps_to_install = $Get_Apps_to_install.Applications.Application

# Show list of apps to install when Hide_Powershell is False
if ($ShowErrors) {
    Write-Host "Applications to install:" -ForegroundColor Yellow
    $appIndex = 1
    foreach ($App in $Apps_to_install) {
        Write-Host "  $appIndex. $($App.Name) - $($App.File)" -ForegroundColor Yellow
        $appIndex++
    }
    Write-Host ""
}

$currentAppIndex = 0
foreach ($App in $Apps_to_install) {
    $currentAppIndex++
    $App_Name = $App.Name
    $App_File = $App.File
    $App_Path = $App.Path
    if ($App_Path) {
        $Folder_Name = $App_Path.split("\")[-1]
        $App_Folder = "$SDBApp_Root_Path\$Folder_Name"
        $App_Full_Path = "$App_Folder\$App_File"
    } else {
        $App_Folder = "$SDBApp_Root_Path"
        $App_Full_Path = "$App_Folder\$App_File"
    }

    $App_CommandLine = $App.CommandLine
    $App_SilentSwitch = $App.Silent_Switch

    # Show what's being processed when Hide_Powershell is False
    if ($ShowErrors) {
        Write-Host ""
        Write-Host "==========================================" -ForegroundColor Green
        Write-Host "[$currentAppIndex/$($Apps_to_install.Count)] Processing: $App_Name" -ForegroundColor Green
        Write-Host "File: $App_Full_Path"
        if ($App_File -like "*.intunewin" -and $App_CommandLine) {
            Write-Host "Command: $App_CommandLine"
        } elseif ($App_SilentSwitch) {
            Write-Host "Arguments: $App_SilentSwitch"
        }
        Write-Host "==========================================" -ForegroundColor Green
    }

    # Check if file exists before trying to execute
    if (-not (Test-Path $App_Full_Path)) {
        Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode -1 -ErrorMessage "File not found"
        continue
    }

    if ( ($App_File -like "*.exe*") -or ($App_File -like "*.msi*") ) {
        try {
            if ($App_SilentSwitch -ne "") {
                $process = Start-Process $App_Full_Path -ArgumentList "$App_SilentSwitch" -Wait -PassThru
            } else {
                $process = Start-Process $App_Full_Path -Wait -PassThru
            }
            
            if ($process.ExitCode -ne 0) {
                Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode $process.ExitCode
            } elseif ($ShowErrors) {
                Write-Host "[SUCCESS] $App_Name installed successfully" -ForegroundColor Green
            }
        } catch {
            Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode -1 -ErrorMessage $_.Exception.Message
        }
    } elseif ( ($App_File -like "*.ps1*") -or ($App_File -like "*.vbs*") ) {
        try {
            & { Invoke-Expression ($App_Full_Path) }
            if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
                Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode $LASTEXITCODE
            }
        } catch {
            Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode -1 -ErrorMessage $_.Exception.Message
        }
    } elseif ($App_File -like "*.intunewin") {
        $Config_Folder_Path = "$Desktop\Intunewin_Config_Folder"
        New-Item -Path $Desktop -Name "Intunewin_Config_Folder" -Type Directory -Force | Out-Null
        $Intunewin_Content_File = "$Config_Folder_Path\Intunewin_Folder.txt"
        $Intunewin_Command_File = "$Config_Folder_Path\Intunewin_Install_Command.txt"

        $App_Full_Path | Out-File $Intunewin_Content_File -Force -NoNewline
        $App_CommandLine | Out-File $Intunewin_Command_File -Force -NoNewline
        
        try {
            # Always use Hidden for the PowerShell wrapper - only the actual installer should be visible
            # Note: IntuneWin_Install.ps1 handles its own error dialogs, so we don't show duplicate errors here
            $process = Start-Process -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
                -ArgumentList "-sta -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -File `"$Sandbox_Root_Path\IntuneWin_Install.ps1`" -Intunewin_Content_File `"$Intunewin_Content_File`" -Intunewin_Command_File `"$Intunewin_Command_File`"" `
                -Wait -PassThru
            
            # Only show success message - errors are handled by IntuneWin_Install.ps1
            if ($process.ExitCode -eq 0 -and $ShowErrors) {
                Write-Host "[SUCCESS] $App_Name installed successfully" -ForegroundColor Green
            } elseif ($process.ExitCode -ne 0 -and $ShowErrors) {
                Write-Host "[FAILED] $App_Name installation failed (exit code: $($process.ExitCode))" -ForegroundColor Red
            }
        } catch {
            Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode -1 -ErrorMessage $_.Exception.Message
        }
    } else {
        try {
            Set-Location $App_Folder
            & { Invoke-Expression (Get-Content -Raw $App_File) }
            & { Invoke-Expression ($App_CommandLine) }
        } catch {
            Show-InstallationError -AppName $App_Name -FilePath $App_Full_Path -ExitCode -1 -ErrorMessage $_.Exception.Message
        }
    }
}