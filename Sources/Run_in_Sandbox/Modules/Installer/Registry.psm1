<#
.SYNOPSIS
    Registry management module for Run-in-Sandbox

.DESCRIPTION
    This module provides registry management functionality for the Run-in-Sandbox application.
    It handles reading, writing, and validating registry entries for installation.
#>

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
    
    $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
    $Shell_Registry_Key = "$Base_Registry_Key\Shell"
    $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
    $MainMenuLabel_Path = "$Shell_Registry_Key\$MainMenuLabel"
    $Command_Path = "$Key_Label_Path\Command"
    $Command_for = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Unrestricted -sta -File C:\\ProgramData\\Run_in_Sandbox\\RunInSandbox.ps1 -Type $Type -ScriptPath `"%V`""
    
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
            if ( -not (Test-Path $MainMenuLabel_Path) ) {
                New-Item -Path $Shell_Registry_Key -Name $MainMenuLabel -Force | Out-Null
                New-ItemProperty -Path $MainMenuLabel_Path -Name "subcommands" -PropertyType String | Out-Null
                New-Item -Path $MainMenuLabel_Path -Name "Shell" -Force | Out-Null
                New-ItemProperty -Path $MainMenuLabel_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -ErrorAction Stop | Out-Null
            }
            $Key_Label_Path = "$MainMenuLabel_Path\Shell\$Key_Label"
            $Command_Path = "$Key_Label_Path\Command"
        }

        if (Test-Path -Path $Key_Label_Path) {
            Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for $Type has already been added"
            Add-Content -Path $RegistryPathsFile -Value $Key_Label_Path
            return
        }

        New-Item -Path $Key_Label_Path -ErrorAction Stop | Out-Null
        New-Item -Path $Command_Path -ErrorAction Stop | Out-Null
        if (-not $MainMenuSwitch) {
            New-ItemProperty -Path $Key_Label_Path -Name "icon" -PropertyType String -Value $Sandbox_Icon -ErrorAction Stop | Out-Null
        }
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
    Write-LogMessage -Message_Type "INFO" -Message "Removing context menu for $Type"
    $Base_Registry_Key = "$Reg_Path\$Sub_Reg_Path"
    $Shell_Registry_Key = "$Base_Registry_Key\Shell"
    $Key_Label_Path = "$Shell_Registry_Key\$Key_Label"
    
    
    if (-not (Test-Path -Path $Key_Label_Path) ) {
        if ($DeepClean) {
            Write-LogMessage -Message_Type "INFO" -Message "Registry Path for $Type has already been removed by deepclean"
            return
        }
        Write-LogMessage -Message_Type "WARNING" -Message "Could not find path for $Type"
        return
    }
    
    try {
        # Get all child items and sort by depth (deepest first)
        $ChildItems = Get-ChildItem -Path $Key_Label_Path -Recurse | Sort-Object { $_.PSPath.Split('\').Count } -Descending

        foreach ($ChildItem in $ChildItems) {
            Remove-Item -LiteralPath $ChildItem.PSPath -Force -ErrorAction Stop
        }

        # Remove the main registry path if it still exists
        if (Test-Path -Path $Key_Label_Path) {
            Remove-Item -LiteralPath $Key_Label_Path -Force -ErrorAction Stop
        }
        
        Write-LogMessage -Message_Type "SUCCESS" -Message "Context menu for `"$Info_Type`" has been removed"
    } catch {
        Write-LogMessage -Message_Type "ERROR" -Message "Context menu for $Type couldn´t be removed"
    }
}

function Find-RegistryIconPaths {
    param (
        [Parameter(Mandatory=$true)] [string]$rootRegistryPath,
        [string]$iconValueToMatch = "C:\\ProgramData\\Run_in_Sandbox\\sandbox.ico"
    )

    # Export the registry at the specified rootRegistryPath
    $exportPath = "$env:TEMP\registry_export.reg"
    reg export $rootRegistryPath $exportPath /y > $null 2>&1

    # Initialize an empty array to store matching paths
    $matchingPaths = @()

    # Read the exported registry file
    $lines = Get-Content -Path $exportPath

    # Process each line in the exported registry file
    foreach ($line in $lines) {
        # Check if the line defines a new key
        if ($line -match '^\[([^\]]+)\]$') {
            $currentPath = $matches[1]
        }

        # If the line contains the icon value, add the current path to the list
        # If the line contains the icon value, add the current path to the list
        if ($line -match '^\s*\"Icon\"=\"([^\"]+)\"$' -and $matches[1] -eq $iconValueToMatch) {
            $currentPath = "REGISTRY::$currentPath"
            $matchingPaths += $currentPath
        }
    }
    $matchingPaths = $matchingPaths | Sort-Object
    return $matchingPaths
}

Export-ModuleMember -Function @(
    'Export-RegConfig',
    'Add-RegItem',
    'Remove-RegItem',
    'Find-RegistryIconPaths'
)