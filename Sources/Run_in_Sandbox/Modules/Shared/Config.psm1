<#
.SYNOPSIS
    Configuration management module for Run-in-Sandbox

.DESCRIPTION
    This module provides configuration management functionality for the Run-in-Sandbox application.
    It handles loading, merging, and managing configuration settings.
#>

# Global variables
[CmdletBinding()] param()

$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"

function Get-Config {
    if ( [string]::IsNullOrEmpty($XML_Config) ) {
        return
    }
    if (-not (Test-Path -Path $XML_Config) ) {
        return
    }
    $Get_XML_Content = [xml](Get-Content $XML_Config)
    
    $Global:Add_EXE = $Get_XML_Content.Configuration.ContextMenu_EXE
    $Global:Add_MSI = $Get_XML_Content.Configuration.ContextMenu_MSI
    $Global:Add_PS1 = $Get_XML_Content.Configuration.ContextMenu_PS1
    $Add_VBS = $Get_XML_Content.Configuration.ContextMenu_VBS
    $Global:Add_ZIP = $Get_XML_Content.Configuration.ContextMenu_ZIP
    $Global:Add_Folder = $Get_XML_Content.Configuration.ContextMenu_Folder
    $Global:Add_Intunewin = $Get_XML_Content.Configuration.ContextMenu_Intunewin
    $Global:Add_MultipleApp = $Get_XML_Content.Configuration.ContextMenu_MultipleApp
    $Global:Add_Reg = $Get_XML_Content.Configuration.ContextMenu_Reg
    $Global:Add_ISO = $Get_XML_Content.Configuration.ContextMenu_ISO
    $Global:Add_PPKG = $Get_XML_Content.Configuration.ContextMenu_PPKG
    $Global:Add_HTML = $Get_XML_Content.Configuration.ContextMenu_HTML
    $Global:Add_MSIX = $Get_XML_Content.Configuration.ContextMenu_MSIX
    $Global:Add_CMD = $Get_XML_Content.Configuration.ContextMenu_CMD
    $Global:Add_PDF = $Get_XML_Content.Configuration.ContextMenu_PDF
}

# Merge user settings from old config into new config
function Merge-SandboxConfig {
    param([string]$OldConfigPath, [string]$NewConfigPath)
    
    try {
        # Validate input files exist
        if (-not (Test-Path $OldConfigPath)) {
            throw "Old config file not found: $OldConfigPath"
        }
        if (-not (Test-Path $NewConfigPath)) {
            throw "New config file not found: $NewConfigPath"
        }
        
        # Load XML with preservation of whitespace and comments
        $OldConfig = New-Object System.Xml.XmlDocument
        $OldConfig.PreserveWhitespace = $true
        $OldConfig.Load($OldConfigPath)
        
        $NewConfig = New-Object System.Xml.XmlDocument
        $NewConfig.PreserveWhitespace = $true
        $NewConfig.Load($NewConfigPath)
        
        # Validate XML structure
        if (-not $OldConfig.Configuration) {
            throw "Invalid old config: missing Configuration element"
        }
        if (-not $NewConfig.Configuration) {
            throw "Invalid new config: missing Configuration element"
        }
        
        # Preserve user settings but skip CurrentVersion (always use new version)
        foreach ($Child in $OldConfig.Configuration.ChildNodes) {
            # Skip non-element nodes
            if ($Child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            if ($Child.Name -eq "CurrentVersion") { continue }
            
            # Find corresponding element in new config
            $NewSetting = $null
            foreach ($node in $NewConfig.Configuration.ChildNodes) {
                if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element -and $node.Name -eq $Child.Name) {
                    $NewSetting = $node
                    break
                }
            }
            
            if ($NewSetting) {
                # Preserve the value with proper data type handling
                if ($Child.InnerText -eq "true" -or $Child.InnerText -eq "false") {
                    # Handle boolean values
                    $NewSetting.InnerText = $Child.InnerText.ToLower()
                } elseif ([string]::IsNullOrEmpty($Child.InnerText)) {
                    # Handle empty values
                    $NewSetting.InnerText = ""
                } else {
                    # Handle regular string values
                    $NewSetting.InnerText = $Child.InnerText
                }
            }
        }
        
        # Save with proper formatting
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Indent = $true
        $settings.IndentChars = "    "
        $settings.NewLineOnAttributes = $false
        $settings.OmitXmlDeclaration = $false
        
        $writer = [System.Xml.XmlWriter]::Create($NewConfigPath, $settings)
        try {
            $NewConfig.Save($writer)
        } finally {
            $writer.Close()
        }
        
        # Validate the saved XML
        $validateConfig = New-Object System.Xml.XmlDocument
        $validateConfig.Load($NewConfigPath)
        
    } catch {
        throw "Error merging configuration: $($_.Exception.Message)"
    }
}

function Merge-ConfigIfNeeded {
    [CmdletBinding()]
    param(
        [bool]$IsInstalled,
        [string]$RunFolder
    )

    if (-not $IsInstalled) { return }

    $ConfigBackup = "$env:TEMP\Sandbox_Config_Backup.xml"
    if (Test-Path "$RunFolder\Sandbox_Config.xml") {
        Copy-Item "$RunFolder\Sandbox_Config.xml" $ConfigBackup -Force
    } else {
        return
    }

    try {
        Write-Info "Merging configuration..." ([ConsoleColor]::Cyan)
        Merge-SandboxConfig -OldConfigPath $ConfigBackup -NewConfigPath "$RunFolder\Sandbox_Config.xml"
    } catch {
        Write-Verbose "Merge-SandboxConfig failed, attempting simple merge: $($_.Exception.Message)"
        try {
            # Load XML with preservation of whitespace and comments
            $oldConfig = New-Object System.Xml.XmlDocument
            $oldConfig.PreserveWhitespace = $true
            $oldConfig.Load($ConfigBackup)
            
            $newConfig = New-Object System.Xml.XmlDocument
            $newConfig.PreserveWhitespace = $true
            $newConfig.Load("$RunFolder\Sandbox_Config.xml")
            
            if ($oldConfig.Configuration -and $newConfig.Configuration) {
                # Preserve user settings but skip CurrentVersion (always use new version)
                foreach ($child in $oldConfig.Configuration.ChildNodes) {
                    # Skip non-element nodes
                    if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                    if ($child.Name -eq "CurrentVersion") { continue }
                    
                    # Find corresponding element in new config
                    $target = $null
                    foreach ($n in $newConfig.Configuration.ChildNodes) {
                        if ($n.NodeType -eq [System.Xml.XmlNodeType]::Element -and $n.Name -eq $child.Name) {
                            $target = $n
                            break
                        }
                    }
                    
                    if ($target) {
                        # Preserve the value with proper data type handling
                        if ($child.InnerText -eq "true" -or $child.InnerText -eq "false") {
                            # Handle boolean values
                            $target.InnerText = $child.InnerText.ToLower()
                        } elseif ([string]::IsNullOrEmpty($child.InnerText)) {
                            # Handle empty values
                            $target.InnerText = ""
                        } else {
                            # Handle regular string values
                            $target.InnerText = $child.InnerText
                        }
                    }
                }
                
                # Save with proper formatting
                $settings = New-Object System.Xml.XmlWriterSettings
                $settings.Indent = $true
                $settings.IndentChars = "    "
                $settings.NewLineOnAttributes = $false
                $settings.OmitXmlDeclaration = $false
                
                $writer = [System.Xml.XmlWriter]::Create("$RunFolder\Sandbox_Config.xml", $settings)
                try {
                    $newConfig.Save($writer)
                } finally {
                    $writer.Close()
                }
                
                # Validate the saved XML
                $validateConfig = New-Object System.Xml.XmlDocument
                $validateConfig.Load("$RunFolder\Sandbox_Config.xml")
            }
        } catch {
            Write-Verbose "Simple merge failed: $($_.Exception.Message)"
        }
    } finally {
        try { Remove-Item $ConfigBackup -Force -ErrorAction SilentlyContinue } catch {}
    }
}

Export-ModuleMember -Function @(
    'Get-Config',
    'Merge-SandboxConfig',
    'Merge-ConfigIfNeeded'
)