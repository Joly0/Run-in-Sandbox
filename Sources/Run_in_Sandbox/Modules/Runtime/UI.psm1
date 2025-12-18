<#
.SYNOPSIS
    User Interface module for Run-in-Sandbox

.DESCRIPTION
    This module provides user interface functionality for Run-in-Sandbox application.
    It handles loading XAML files and common UI assembly loading.
#>

# Load required assemblies for WPF dialogs
function Initialize-UIAssemblies {
    param(
        [string]$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    )
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
}

# Consolidated function to load XML documents
function Import-XmlDocument {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    
    if (-not (Test-Path $Path)) {
        throw "XAML file not found: $Path"
    }
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($Path)
    return $XamlLoader
}

# Create alias for backward compatibility
Set-Alias -Name LoadXml -Value Import-XmlDocument

Export-ModuleMember -Function @(
    'Initialize-UIAssemblies',
    'Import-XmlDocument'
) -Alias @(
    'LoadXml'
)
