<#
.SYNOPSIS
    Registry management module for Run-in-Sandbox

.DESCRIPTION
    This module provides registry management functionality for the Run-in-Sandbox application.
    It handles reading, writing, and validating registry entries for installation.
#>
[CmdletBinding()] param()

# Import Environment module to access $Sandbox_Icon variable
If (Test-Path -LiteralPath "$PSScriptRoot\..\Shared\Environment.psm1") {
    Import-Module "$PSScriptRoot\..\Shared\Environment.psm1" -Force
}

# Function to export registry configuration
function Export-RegConfig {
    param (
        [string] $Reg_Path,
        [string] $Backup_Folder = "$Run_in_Sandbox_Folder\Registry_Backup",
        [string] $Type,
        [string] $Sub_Reg_Path
    )
    
    if ($Exported_Keys -contains $Reg_Path) {
        $Exported_Keys.Add($Reg_Path)
    } else {
        return
    }
    
    if (-not (Test-Path $Backup_Folder) ) {
        New-Item -ItemType Directory -Path $Backup_Folder -Force | Out-Null
    }
    
    Write-LogMessage -Message_Type "INFO" -Message "Exporting registry keys"
    
    $Backup_Path = $Backup_Folder + "\" + "Backup_" + $Type
    if ($Sub_Reg_Path) {
        $Backup_Path = $Backup_Path + "_" + $Sub_Reg_Path
    }
    $Backup_Path = $Backup_Path + ".reg"
    
    reg export $Reg_Path $Backup_Path /y > $null 2>&1

    # Check if the command ran successfully
    if ($?) {
        Write-LogMessage -Message_Type "SUCCESS" -Message "Exported `"$Reg_Path`" to `"$Backup_Path`""
    } else {
        Write-LogMessage -Message_Type "ERROR" -Message "Failed to export `"$Reg_Path`""
    }
}

# Function to test if a registry entry is complete
function Test-RegistryEntryComplete {
    param (
        [Parameter(Mandatory=$true)] [string] $Key_Label_Path,
        [Parameter(Mandatory=$true)] [string] $Command_Path,
        [Parameter(Mandatory=$true)] [string] $Type,
        [string] $Icon_Path,
        [switch] $MainMenuSwitch
    )
    
    Write-Verbose "Test-RegistryEntryComplete: Starting validation for $Type"
    Write-Verbose "Test-RegistryEntryComplete: Key_Label_Path = $Key_Label_Path"
    Write-Verbose "Test-RegistryEntryComplete: Command_Path = $Command_Path"
    Write-Verbose "Test-RegistryEntryComplete: Icon_Path = $Icon_Path"
    Write-Verbose "Test-RegistryEntryComplete: MainMenuSwitch = $MainMenuSwitch"
    
    # Check if the main registry key exists
    if (-not (Test-Path -Path $Key_Label_Path)) {
        Write-Verbose "Test-RegistryEntryComplete: Registry key does not exist: $Key_Label_Path"
        return $false
    }
    
    # Check if the command subkey exists
    if (-not (Test-Path -Path $Command_Path)) {
        Write-Verbose "Test-RegistryEntryComplete: Command subkey does not exist: $Command_Path"
        return $false
    }
    
    # Check if the command value is correct
    $Expected_Command = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
    try {
        $Actual_Command = (Get-Item -Path $Command_Path).Value
        if ($Actual_Command -ne $Expected_Command) {
            Write-Verbose "Test-RegistryEntryComplete: Command value mismatch"
            Write-Verbose "Test-RegistryEntryComplete: Expected: $Expected_Command"
            Write-Verbose "Test-RegistryEntryComplete: Actual: $Actual_Command"
            return $false
        }
    } catch {
        Write-Verbose "Test-RegistryEntryComplete: Failed to read command value: $($_.Exception.Message)"
        return $false
    }
    
    # Check if the icon property exists and has the correct value
    if ($Icon_Path) {
        try {
            $Icon_Property = Get-ItemProperty -Path $Icon_Path -Name "icon" -ErrorAction SilentlyContinue
            if (-not $Icon_Property) {
                Write-Verbose "Test-RegistryEntryComplete: Icon property does not exist at: $Icon_Path"
                return $false
            }
            
            # Ensure we have access to the Sandbox_Icon variable
            if (-not $Global:Sandbox_Icon) {
                $Global:Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
            }
            
            if ($Icon_Property.icon -ne $Global:Sandbox_Icon) {
                Write-Verbose "Test-RegistryEntryComplete: Icon value mismatch"
                Write-Verbose "Test-RegistryEntryComplete: Expected: $Global:Sandbox_Icon"
                Write-Verbose "Test-RegistryEntryComplete: Actual: $($Icon_Property.icon)"
                return $false
            }
        } catch {
            Write-Verbose "Test-RegistryEntryComplete: Failed to read icon property: $($_.Exception.Message)"
            return $false
        }
    }
    
    Write-Verbose "Test-RegistryEntryComplete: Registry entry is complete"
    return $true
}

# Function to add a registry item
function Add-RegItem {
    param (
        [string] $Reg_Path = "Registry::HKEY_CLASSES_ROOT",
        [string] $Sub_Reg_Path,
        [string] $Type,
        [string] $Entry_Name = $Type,
        [string] $Info_Type = $Type,
        [string] $Key_Label = "Run $Entry_Name in Sandbox",
        [string] $RegistryPathsFile = "$Run_in_Sandbox_Folder\RegistryEntries.txt",
        [string] $MainMenuLabel,
        [switch] $MainMenuSwitch
    )
    
    # Verbose logging for debugging
    Write-Verbose "Add-RegItem: Starting registry item addition"
    Write-Verbose "Add-RegItem: Type = $Type, Entry_Name = $Entry_Name"
    Write-Verbose "Add-RegItem: Reg_Path = $Reg_Path"
    Write-Verbose "Add-RegItem: Sub_Reg_Path = $Sub_Reg_Path"
    Write-Verbose "Add-RegItem: Key_Label = $Key_Label"
    Write-Verbose "Add-RegItem: MainMenuSwitch = $MainMenuSwitch"
    # Ensure we have access to the Sandbox_Icon variable
    if (-not $Global:Sandbox_Icon) {
        $Global:Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
    }
    Write-Verbose "Add-RegItem: Sandbox_Icon variable value = $Global:Sandbox_Icon"
    Write-Verbose "Add-RegItem: Sandbox_Icon file exists = $(Test-Path $Global:Sandbox_Icon)"
    
    $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
    $Shell_Registry_Key = "$Base_Registry_Key\Shell"
    $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
    $MainMenuLabel_Path = "$Shell_Registry_Key\$MainMenuLabel"
    $Command_Path = "$Key_Label_Path\Command"
    $Command_for = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
    
    Write-Verbose "Add-RegItem: Base_Registry_Key = $Base_Registry_Key"
    Write-Verbose "Add-RegItem: Shell_Registry_Key = $Shell_Registry_Key"
    Write-Verbose "Add-RegItem: Key_Label_Path = $Key_Label_Path"
    
    Export-RegConfig -Reg_Path $($Base_Registry_Key.Split("::")[-1]) -Type $Type -Sub_Reg_Path $Sub_Reg_Path -ErrorAction Continue
    
    try {
        # Log the root registry path to the specified file
        if (-not (Test-Path $RegistryPathsFile) ) {
            New-Item -ItemType File -Path $RegistryPathsFile -Force | Out-Null
        }

        if (-not (Test-Path -Path $Base_Registry_Key) ) {
            New-Item -Path $Base_Registry_Key -ErrorAction Stop | Out-Null
        }
        
        if (-not (Test-Path -Path $Shell_Registry_Key) ) {
            New-Item -Path $Shell_Registry_Key -ErrorAction Stop | Out-Null
        }
        
        if ($MainMenuSwitch) {
            Write-Verbose "Add-RegItem: Processing MainMenuSwitch"
            if ( -not (Test-Path $MainMenuLabel_Path) ) {
                Write-Verbose "Add-RegItem: Creating MainMenuLabel_Path = $MainMenuLabel_Path"
                New-Item -Path $Shell_Registry_Key -Name $MainMenuLabel -Force | Out-Null
                New-ItemProperty -Path $MainMenuLabel_Path -Name "subcommands" -PropertyType String | Out-Null
                New-Item -Path $MainMenuLabel_Path -Name "Shell" -Force | Out-Null
                Write-Verbose "Add-RegItem: Setting icon property at $MainMenuLabel_Path with value $Global:Sandbox_Icon"
                try {
                    New-ItemProperty -Path $MainMenuLabel_Path -Name "icon" -PropertyType String -Value $Global:Sandbox_Icon -ErrorAction Stop | Out-Null
                    Write-Verbose "Add-RegItem: Successfully set icon property for main menu"
                } catch {
                    Write-Verbose "Add-RegItem: Failed to set icon property for main menu: $($_.Exception.Message)"
                    throw
                }
            } else {
                Write-Verbose "Add-RegItem: MainMenuLabel_Path already exists, skipping creation"
            }
            $Key_Label_Path = "$MainMenuLabel_Path\Shell\$Key_Label"
            $Command_Path = "$Key_Label_Path\Command"
            Write-Verbose "Add-RegItem: Updated Key_Label_Path for MainMenuSwitch = $Key_Label_Path"
        }

        # Determine the correct icon path based on MainMenuSwitch
        $Icon_Path = if ($MainMenuSwitch) { $MainMenuLabel_Path } else { $Key_Label_Path }
        
        # Check if the registry entry exists and is complete
        $Is_Entry_Complete = Test-RegistryEntryComplete -Key_Label_Path $Key_Label_Path -Command_Path $Command_Path -Type $Type -Icon_Path $Icon_Path -MainMenuSwitch:$MainMenuSwitch
        
        if ($Is_Entry_Complete) {
            Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for $Type has already been added and is complete"
            Add-Content -Path $RegistryPathsFile -Value $Key_Label_Path
            return
        }
        
        # If the registry key exists but is incomplete, repair it
        if (Test-Path -Path $Key_Label_Path) {
            Write-LogMessage -Message_Type "INFO" -Message "Context menu for $Type exists but is incomplete, attempting repair"
            
            # Repair the command subkey if needed
            if (-not (Test-Path -Path $Command_Path)) {
                Write-Verbose "Add-RegItem: Repairing missing command subkey: $Command_Path"
                try {
                    New-Item -Path $Command_Path -ErrorAction Stop | Out-Null
                    $Expected_Command = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
                    Set-Item -Path $Command_Path -Value $Expected_Command -Force -ErrorAction Stop | Out-Null
                    Write-Verbose "Add-RegItem: Successfully repaired command subkey"
                } catch {
                    Write-Verbose "Add-RegItem: Failed to repair command subkey: $($_.Exception.Message)"
                    Write-LogMessage -Message_Type "ERROR" -Message "Failed to repair command subkey for $Type"
                }
            } else {
                # Check and repair the command value if needed
                $Expected_Command = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
                try {
                    $Actual_Command = (Get-Item -Path $Command_Path).Value
                    if ($Actual_Command -ne $Expected_Command) {
                        Write-Verbose "Add-RegItem: Repairing incorrect command value"
                        Set-Item -Path $Command_Path -Value $Expected_Command -Force -ErrorAction Stop | Out-Null
                        Write-Verbose "Add-RegItem: Successfully repaired command value"
                    }
                } catch {
                    Write-Verbose "Add-RegItem: Failed to repair command value: $($_.Exception.Message)"
                }
            }
            
            # Repair the icon property if needed
            if ($Icon_Path) {
                try {
                    $Icon_Property = Get-ItemProperty -Path $Icon_Path -Name "icon" -ErrorAction SilentlyContinue
                    if (-not $Icon_Property -or $Icon_Property.icon -ne $Global:Sandbox_Icon) {
                        Write-Verbose "Add-RegItem: Repairing icon property at $Icon_Path with value $Global:Sandbox_Icon"
                        New-ItemProperty -Path $Icon_Path -Name "icon" -PropertyType String -Value $Global:Sandbox_Icon -Force -ErrorAction Stop | Out-Null
                        Write-Verbose "Add-RegItem: Successfully repaired icon property"
                    }
                } catch {
                    Write-Verbose "Add-RegItem: Failed to repair icon property: $($_.Exception.Message)"
                }
            }
            
            # Check if the entry is now complete after repair
            $Is_Entry_Complete_After_Repair = Test-RegistryEntryComplete -Key_Label_Path $Key_Label_Path -Command_Path $Command_Path -Type $Type -Icon_Path $Icon_Path -MainMenuSwitch:$MainMenuSwitch
            
            if ($Is_Entry_Complete_After_Repair) {
                Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for $Type has been successfully repaired"
                Add-Content -Path $RegistryPathsFile -Value $Key_Label_Path
                return
            } else {
                Write-LogMessage -Message_Type "WARNING" -Message "Context menu for $Type could not be fully repaired, proceeding with full recreation"
            }
        }

        Write-Verbose "Add-RegItem: Creating registry key: $Key_Label_Path"
        New-Item -Path $Key_Label_Path -ErrorAction Stop | Out-Null
        Write-Verbose "Add-RegItem: Creating command path: $Command_Path"
        New-Item -Path $Command_Path -ErrorAction Stop | Out-Null
        if (-not $MainMenuSwitch) {
            Write-Verbose "Add-RegItem: Setting icon property at $Key_Label_Path with value $Global:Sandbox_Icon"
            try {
                New-ItemProperty -Path $Key_Label_Path -Name "icon" -PropertyType String -Value $Global:Sandbox_Icon -ErrorAction Stop | Out-Null
                Write-Verbose "Add-RegItem: Successfully set icon property"
            } catch {
                Write-Verbose "Add-RegItem: Failed to set icon property: $($_.Exception.Message)"
                throw
            }
        }
        Write-Verbose "Add-RegItem: Setting command value at $Command_Path"
        Set-Item -Path $Command_Path -Value $Command_for -Force -ErrorAction Stop | Out-Null

        Add-Content -Path $RegistryPathsFile -Value $Key_Label_Path

        Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been added"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Context menu for $Type could not be added"
    }
}

# Function to remove a registry item
function Remove-RegItem {
    param (
        [string] $Reg_Path = "Registry::HKEY_CLASSES_ROOT",
        [Parameter(Mandatory=$true)] [string] $Sub_Reg_Path,
        [Parameter(Mandatory=$true)] [string] $Type,
        [string] $Entry_Name = $Type,
        [string] $Info_Type = $Type,
        [string] $Key_Label = "Run $Entry_Name in Sandbox",
        [string] $MainMenuLabel,
        [switch] $MainMenuSwitch
    )
    
    # Verbose logging for debugging
    Write-Verbose "Remove-RegItem: Starting registry item removal"
    Write-Verbose "Remove-RegItem: Type = $Type, Entry_Name = $Entry_Name"
    Write-Verbose "Remove-RegItem: Reg_Path = $Reg_Path"
    Write-Verbose "Remove-RegItem: Sub_Reg_Path = $Sub_Reg_Path"
    Write-Verbose "Remove-RegItem: Key_Label = $Key_Label"
    Write-Verbose "Remove-RegItem: MainMenuSwitch = $MainMenuSwitch"
    Write-Verbose "Remove-RegItem: Global:DeepClean = $Global:DeepClean"
    
    Write-LogMessage -Message_Type "INFO" -Message "Removing context menu for $Type"
    $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
    $Shell_Registry_Key = "$Base_Registry_Key\Shell"
    $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
    
    Write-Verbose "Remove-RegItem: Base_Registry_Key = $Base_Registry_Key"
    Write-Verbose "Remove-RegItem: Shell_Registry_Key = $Shell_Registry_Key"
    Write-Verbose "Remove-RegItem: Key_Label_Path = $Key_Label_Path"
    
    if (-not (Test-Path -Path $Key_Label_Path) ) {
        Write-Verbose "Remove-RegItem: Registry path does not exist: $Key_Label_Path"
        if ($Global:DeepClean) {
            Write-Verbose "Remove-RegItem: DeepClean mode enabled, path already removed"
            Write-LogMessage -Message_Type "INFO" -Message "Registry Path for $Type has already been removed by deepclean"
            return
        }
        Write-Verbose "Remove-RegItem: Path not found and DeepClean is not enabled"
        Write-LogMessage -Message_Type "WARNING" -Message "Could not find path for $Type"
        return
    }
    
    Write-Verbose "Remove-RegItem: Registry path exists, proceeding with removal"
    
    try {
        # Get all child items and sort by depth (deepest first)
        Write-Verbose "Remove-RegItem: Getting child items for recursive removal"
        $ChildItems = Get-ChildItem -Path $Key_Label_Path -Recurse | Sort-Object { $_.PSPath.Split('\').Count } -Descending
        Write-Verbose "Remove-RegItem: Found $($ChildItems.Count) child items to remove"

        foreach ($ChildItem in $ChildItems) {
            Write-Verbose "Remove-RegItem: Removing child item: $($ChildItem.PSPath)"
            Remove-Item -LiteralPath $ChildItem.PSPath -Force -ErrorAction Stop
        }

        # Remove the main registry path if it still exists
        if (Test-Path -Path $Key_Label_Path) {
            Write-Verbose "Remove-RegItem: Removing main registry path: $Key_Label_Path"
            Remove-Item -LiteralPath $Key_Label_Path -Force -ErrorAction Stop
        } else {
            Write-Verbose "Remove-RegItem: Main registry path already removed"
        }
        
        Write-Verbose "Remove-RegItem: Successfully removed registry items"
        Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been removed"
    } catch {
        Write-Verbose "Remove-RegItem: Failed to remove registry items: $($_.Exception.Message)"
        Write-LogMessage -Message_Type "ERROR" -Message "Context menu for $Type couldn´t be removed"
    }
}

function Find-RegistryIconPaths {
    param (
        [Parameter(Mandatory=$true)] [string]$rootRegistryPath,
        [string]$iconValueToMatch = "C:\\ProgramData\\Run_in_Sandbox\\sandbox.ico"
    )

    Write-Verbose "Find-RegistryIconPaths: Starting search in $rootRegistryPath"
    Write-Verbose "Find-RegistryIconPaths: Looking for icon value: $iconValueToMatch"

    # Export the registry at the specified rootRegistryPath
    $exportPath = "$env:TEMP\registry_export_$($rootRegistryPath.Replace('\', '_').Replace(':', '')).reg"
    Write-Verbose "Find-RegistryIconPaths: Exporting registry to $exportPath"
    
    try {
        reg export $rootRegistryPath $exportPath /y > $null 2>&1
        Write-Verbose "Find-RegistryIconPaths: Registry export completed successfully"
    } catch {
        Write-Verbose "Find-RegistryIconPaths: Registry export failed: $($_.Exception.Message)"
        return @()
    }

    # Initialize an empty array to store matching paths
    $matchingPaths = @()

    # Read the exported registry file
    try {
        $lines = Get-Content -Path $exportPath
        Write-Verbose "Find-RegistryIconPaths: Read $($lines.Count) lines from export file"
    } catch {
        Write-Verbose "Find-RegistryIconPaths: Failed to read export file: $($_.Exception.Message)"
        return @()
    }

    # Process each line in the exported registry file
    $currentPath = $null
    foreach ($line in $lines) {
        # Check if the line defines a new key
        if ($line -match '^\[([^\]]+)\]$') {
            $currentPath = $matches[1]
            Write-Verbose "Find-RegistryIconPaths: Found registry key: $currentPath"
        }

        # If the line contains the icon value, add the current path to the list
        if ($line -match '^\s*\"Icon\"=\"([^\"]+)\"$' -and $matches[1] -eq $iconValueToMatch) {
            $registryPath = "REGISTRY::$currentPath"
            $matchingPaths += $registryPath
            Write-Verbose "Find-RegistryIconPaths: Found matching icon at: $registryPath"
        }
    }
    
    Write-Verbose "Find-RegistryIconPaths: Found $($matchingPaths.Count) total matching paths"
    $matchingPaths = $matchingPaths | Sort-Object
    
    # Clean up temporary file
    try {
        Remove-Item $exportPath -Force -ErrorAction SilentlyContinue
        Write-Verbose "Find-RegistryIconPaths: Cleaned up temporary export file"
    } catch {
        Write-Verbose "Find-RegistryIconPaths: Failed to clean up temporary file: $($_.Exception.Message)"
    }
    
    return $matchingPaths
}

Export-ModuleMember -Function @(
    'Export-RegConfig',
    'Add-RegItem',
    'Remove-RegItem',
    'Find-RegistryIconPaths',
    'Test-RegistryEntryComplete'
)