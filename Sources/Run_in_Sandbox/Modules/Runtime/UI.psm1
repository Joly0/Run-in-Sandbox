<#
.SYNOPSIS
    User Interface module for Run-in-Sandbox

.DESCRIPTION
    This module provides user interface functionality for Run-in-Sandbox application.
    It handles dialog boxes, notifications, and user interaction elements.
#>

# Consolidated function to load XML documents
function Import-XmlDocument {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($Path)
    return $XamlLoader
}

# Create alias for backward compatibility
Set-Alias -Name LoadXml -Value Import-XmlDocument

Export-ModuleMember -Function @(
    'Import-XmlDocument'
) -Alias @(
    'LoadXml'
)