<#
.SYNOPSIS
    Logging module for Run-in-Sandbox

.DESCRIPTION
    This module provides logging functionality for the Run-in-Sandbox application.
    It handles writing log messages to files and displaying information to users.
#>

# Define global variables
$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"

# Function to write log messages
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Message_Type
    )

    $MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    Add-Content -Path $Log_File -Value "$MyDate - $Message_Type : $Message"
    $ForegroundColor = switch ($Message_Type) {
        "INFO"    { 'White' }
        "SUCCESS" { 'Green' }
        "WARNING" { 'Yellow' }
        "ERROR"   { 'DarkRed' }
        default   { 'White' }
    }
    Write-Host "$MyDate - $Message_Type : $Message" -ForegroundColor $ForegroundColor
}

# Function to write info messages (alias to Write-LogMessage)
function Write-Info {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    # Only show to end-user for relevant messages (not too chatty)
    Write-Host $Message -ForegroundColor $Color
}

Export-ModuleMember -Function @(
    'Write-LogMessage',
    'Write-Info'
)