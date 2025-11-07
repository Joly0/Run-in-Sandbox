<#
.SYNOPSIS
    Validation module for Run-in-Sandbox

.DESCRIPTION
    This module provides validation functionality for the Run-in-Sandbox application.
    It handles validation of installation requirements, files, and system state.
#>

function Validate-Installation {
    param([string]$RunFolder)

    $RequiredFiles = @(
        (Join-Path $RunFolder "RunInSandbox.ps1"),
        (Join-Path $RunFolder "CommonFunctions.ps1"),
        (Join-Path $RunFolder "Sandbox_Config.xml"),
        (Join-Path $RunFolder "version.json")
    )
    
    # Check for the Modules folder structure
    $RequiredFolders = @(
        (Join-Path $RunFolder "Modules"),
        (Join-Path $RunFolder "Modules\Shared"),
        (Join-Path $RunFolder "Modules\Installer"),
        (Join-Path $RunFolder "Modules\Runtime"),
        (Join-Path $RunFolder "startup-scripts")
    )
    
    # Check for key module files
    $RequiredModuleFiles = @(
        (Join-Path $RunFolder "Modules\Shared\Logging.psm1"),
        (Join-Path $RunFolder "Modules\Shared\Config.psm1"),
        (Join-Path $RunFolder "Modules\Installer\Core.psm1"),
        (Join-Path $RunFolder "Modules\Runtime\Dialogs.psm1")
    )

    $ok = $true
    foreach ($f in $RequiredFiles) {
        if (-not (Test-Path $f)) {
            Write-Info "Validation failed: Missing $f" ([ConsoleColor]::Red)
            $ok = $false
        }
    }
    
    foreach ($folder in $RequiredFolders) {
        if (-not (Test-Path $folder)) {
            Write-Info "Validation failed: Missing folder $folder" ([ConsoleColor]::Red)
            $ok = $false
        }
    }
    
    foreach ($f in $RequiredModuleFiles) {
        if (-not (Test-Path $f)) {
            Write-Info "Validation failed: Missing module file $f" ([ConsoleColor]::Red)
            $ok = $false
        }
    }
    
    return $ok
}

Export-ModuleMember -Function @(
    'Validate-Installation'
)