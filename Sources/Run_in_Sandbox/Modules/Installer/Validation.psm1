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

    $ok = $true
    foreach ($f in $RequiredFiles) {
        if (-not (Test-Path $f)) {
            Write-Info "Validation failed: Missing $f" ([ConsoleColor]::Red)
            $ok = $false
        }
    }
    return $ok
}

Export-ModuleMember -Function @(
    'Validate-Installation'
)