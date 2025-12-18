<#
.SYNOPSIS
    Configuration management module for Run-in-Sandbox

.DESCRIPTION
    This module provides configuration management functionality for the Run-in-Sandbox application.
    It handles loading and managing configuration settings.

.NOTES
    This module expects Environment.psm1 to be loaded first to provide global variables.
    If Environment.psm1 is not loaded, it will use default paths.
#>

# Ensure global variables are set (fallback if Environment.psm1 not loaded)
if (-not $Global:Run_in_Sandbox_Folder) {
    $Global:Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
}
if (-not $Global:XML_Config) {
    $Global:XML_Config = "$Global:Run_in_Sandbox_Folder\Sandbox_Config.xml"
}

# Reads Sandbox_Config.xml and populates global variables for context menu settings
function Get-Config {
    if ( [string]::IsNullOrEmpty($Global:XML_Config) ) {
        return
    }
    if (-not (Test-Path -Path $Global:XML_Config) ) {
        return
    }
    $Get_XML_Content = [xml](Get-Content $Global:XML_Config)
    
    $Global:Add_EXE = $Get_XML_Content.Configuration.ContextMenu_EXE
    $Global:Add_MSI = $Get_XML_Content.Configuration.ContextMenu_MSI
    $Global:Add_PS1 = $Get_XML_Content.Configuration.ContextMenu_PS1
    $Global:Add_VBS = $Get_XML_Content.Configuration.ContextMenu_VBS
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

# Merges user settings from old config into new config while preserving new config structure
function Merge-SandboxConfig {
    param(
        [string]$OldConfigPath,
        [string]$NewConfigPath
    )
    
    try {
        if (-not (Test-Path $OldConfigPath)) {
            throw "Old config file not found: $OldConfigPath"
        }
        if (-not (Test-Path $NewConfigPath)) {
            throw "New config file not found: $NewConfigPath"
        }
        
        # Load XML with preservation of whitespace
        $OldConfig = New-Object System.Xml.XmlDocument
        $OldConfig.PreserveWhitespace = $true
        $OldConfig.Load($OldConfigPath)
        
        $NewConfig = New-Object System.Xml.XmlDocument
        $NewConfig.PreserveWhitespace = $true
        $NewConfig.Load($NewConfigPath)
        
        if (-not $OldConfig.Configuration) {
            throw "Invalid old config: missing Configuration element"
        }
        if (-not $NewConfig.Configuration) {
            throw "Invalid new config: missing Configuration element"
        }
        
        # Preserve user settings but skip CurrentVersion (always use new version)
        foreach ($Child in $OldConfig.Configuration.ChildNodes) {
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
                # Preserve the original value exactly as-is (including casing)
                if ([string]::IsNullOrEmpty($Child.InnerText)) {
                    $NewSetting.InnerText = ""
                } else {
                    $NewSetting.InnerText = $Child.InnerText
                }
            }
        }
        
        # Save without XML declaration to match original format
        $NewConfig.Save($NewConfigPath)
        
        # Validate the saved XML
        $validateConfig = New-Object System.Xml.XmlDocument
        $validateConfig.Load($NewConfigPath)
        
    } catch {
        throw "Error merging configuration: $($_.Exception.Message)"
    }
}

# Merges config if this is an update (preserves user settings while adding new options)
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
        Write-Verbose "Merging configuration..."
        Merge-SandboxConfig -OldConfigPath $ConfigBackup -NewConfigPath "$RunFolder\Sandbox_Config.xml"
        Write-Verbose "Configuration merged successfully"
    } catch {
        Write-Verbose "Merge-SandboxConfig failed, attempting simple merge: $($_.Exception.Message)"
        try {
            $oldConfig = New-Object System.Xml.XmlDocument
            $oldConfig.PreserveWhitespace = $true
            $oldConfig.Load($ConfigBackup)
            
            $newConfig = New-Object System.Xml.XmlDocument
            $newConfig.PreserveWhitespace = $true
            $newConfig.Load("$RunFolder\Sandbox_Config.xml")
            
            if ($oldConfig.Configuration -and $newConfig.Configuration) {
                foreach ($child in $oldConfig.Configuration.ChildNodes) {
                    if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                    if ($child.Name -eq "CurrentVersion") { continue }
                    
                    $target = $null
                    foreach ($n in $newConfig.Configuration.ChildNodes) {
                        if ($n.NodeType -eq [System.Xml.XmlNodeType]::Element -and $n.Name -eq $child.Name) {
                            $target = $n
                            break
                        }
                    }
                    
                    if ($target) {
                        # Preserve the original value exactly as-is (including casing)
                        if ([string]::IsNullOrEmpty($child.InnerText)) {
                            $target.InnerText = ""
                        } else {
                            $target.InnerText = $child.InnerText
                        }
                    }
                }
                
                # Save without XML declaration to match original format
                $newConfig.Save("$RunFolder\Sandbox_Config.xml")
                
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
