<#
.SYNOPSIS
    Logging module for Run-in-Sandbox

.DESCRIPTION
    This module provides logging functionality for the Run-in-Sandbox application.
    It handles writing log messages to files and displaying information to users.

.NOTES
    This module expects Environment.psm1 to be loaded first to provide $Global:Log_File.
    If Environment.psm1 is not loaded, it will use a default log file path.
#>

# Ensure Log_File is set (fallback if Environment.psm1 not loaded)
if (-not $Global:Log_File) {
    $Global:Log_File = "$env:temp\RunInSandbox_Install.log"
}

# Writes a timestamped message to console (with color coding) and log file
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Message_Type
    )

    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Add-Content -Path $Global:Log_File -Value "$MyDate - $Message_Type : $Message"
    $ForegroundColor = switch ($Message_Type) {
        "INFO"    { 'White' }
        "SUCCESS" { 'Green' }
        "WARNING" { 'Yellow' }
        "ERROR"   { 'DarkRed' }
        default   { 'White' }
    }
    Write-Host "$MyDate - $Message_Type : $Message" -ForegroundColor $ForegroundColor
}

# Writes a detailed message to log file only (not displayed to user)
# Use this for detailed registry paths, debug info, etc.
function Write-LogDetail {
    param (
        [string]$Message,
        [string]$Message_Type = "DETAIL"
    )

    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Add-Content -Path $Global:Log_File -Value "$MyDate - $Message_Type : $Message"
}

Export-ModuleMember -Function @('Write-LogMessage', 'Write-LogDetail')
