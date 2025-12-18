<#
.SYNOPSIS
    Dialog functions module for Run-in-Sandbox

.DESCRIPTION
    This module provides dialog functionality for the Run-in-Sandbox application.
    It handles displaying various dialog boxes for user interaction.
#>

# Show EXE dialog and return the startup command
function Show-EXEDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Run_in_Sandbox_Folder,
        [Parameter(Mandatory=$true)]
        [string]$Full_Startup_Path_Quoted,
        [Parameter(Mandatory=$true)]
        [string]$PSRun_File
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_EXE.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($XamlPath)
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlLoader)
    $Form_EXE = [Windows.Markup.XamlReader]::Load($Reader)
    $EXE_Command_File = "$Run_in_Sandbox_Folder\EXE_Command_File.txt"

    $switches_for_exe = $Form_EXE.findname("switches_for_exe")
    $add_switches = $Form_EXE.findname("add_switches")

    $Script:DialogStartupCommand = ""
    
    $add_switches.Add_Click({
        $Script:Switches_EXE = $switches_for_exe.Text.ToString()
        $Script:DialogStartupCommand = $Full_Startup_Path_Quoted + " " + $Script:Switches_EXE
        $Script:DialogStartupCommand | Out-File $EXE_Command_File -Force -NoNewline
        $Form_EXE.close()
    })

    $Form_EXE.Add_Closing({
        $Script:Switches_EXE = $switches_for_exe.Text.ToString()
        $Script:DialogStartupCommand = $Full_Startup_Path_Quoted + " " + $Script:Switches_EXE
        $Script:DialogStartupCommand | Out-File $EXE_Command_File -Force -NoNewline
    })

    $Form_EXE.ShowDialog() | Out-Null

    # Return the command to run the EXE installer script
    $Sandbox_Root_Path = "C:\Run_in_Sandbox"
    $EXE_Installer = "$Sandbox_Root_Path\EXE_Install.ps1"
    return ($PSRun_File + " " + "$EXE_Installer")
}

# Show MSI dialog and return the startup command
function Show-MSIDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Run_in_Sandbox_Folder,
        [Parameter(Mandatory=$true)]
        [string]$Full_Startup_Path_UnQuoted
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Config.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($XamlPath)
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlLoader)
    $Form_MSI = [Windows.Markup.XamlReader]::Load($Reader)

    $switches_for_exe = $Form_MSI.findname("switches_for_exe")
    $add_switches = $Form_MSI.findname("add_switches")

    $Script:DialogStartupCommand = ""
    
    $add_switches.Add_Click({
        $Script:Switches_MSI = $switches_for_exe.Text.ToString()
        $Script:DialogStartupCommand = "msiexec /i `"$Full_Startup_Path_UnQuoted`" " + $Script:Switches_MSI
        $Form_MSI.close()
    })

    $Form_MSI.Add_Closing({
        $Script:Switches_MSI = $switches_for_exe.Text.ToString()
        $Script:DialogStartupCommand = "msiexec /i `"$Full_Startup_Path_UnQuoted`" " + $Script:Switches_MSI
    })

    $Form_MSI.ShowDialog() | Out-Null
    
    return $Script:DialogStartupCommand
}


# Show PS1 Params dialog and return the startup command
function Show-PS1ParamsDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Run_in_Sandbox_Folder,
        [Parameter(Mandatory=$true)]
        [string]$Full_Startup_Path_UnQuoted,
        [Parameter(Mandatory=$true)]
        [string]$PSRun_File
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($XamlPath)
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlLoader)
    $Form_PS1 = [Windows.Markup.XamlReader]::Load($Reader)

    $parameters_to_add = $Form_PS1.findname("parameters_to_add")
    $add_parameters = $Form_PS1.findname("add_parameters")

    $Script:DialogStartupCommand = ""
    
    $add_parameters.add_click({
        $Script:Paramaters = $parameters_to_add.Text.ToString()
        $Script:DialogStartupCommand = $PSRun_File + " " + "`"$Full_Startup_Path_UnQuoted`" " + $Script:Paramaters
        $Form_PS1.close()
    })

    $Form_PS1.Add_Closing({
        $Script:Paramaters = $parameters_to_add.Text.ToString()
        $Script:DialogStartupCommand = $PSRun_File + " " + "`"$Full_Startup_Path_UnQuoted`" " + $Script:Paramaters
    })

    $Form_PS1.ShowDialog() | Out-Null
    
    return $Script:DialogStartupCommand
}

# Show Intunewin dialog and return the startup command
function Show-IntunewinDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Run_in_Sandbox_Folder,
        [Parameter(Mandatory=$true)]
        [string]$FileName,
        [Parameter(Mandatory=$true)]
        [string]$PSRun_File
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Intunewin.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    $Intunewin_Folder = "C:\IntuneWin\$FileName.intunewin"
    $Intunewin_Content_File = "$Run_in_Sandbox_Folder\Intunewin_Folder.txt"
    $Intunewin_Command_File = "$Run_in_Sandbox_Folder\Intunewin_Install_Command.txt"
    $Intunewin_Folder | Out-File $Intunewin_Content_File -Force -NoNewline
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($XamlPath)
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlLoader)
    $Form_Intunewin = [Windows.Markup.XamlReader]::Load($Reader)

    $install_command_intunewin = $Form_Intunewin.findname("install_command_intunewin")
    $add_install_command = $Form_Intunewin.findname("add_install_command")

    $add_install_command.add_click({
        $Script:install_command = $install_command_intunewin.Text.ToString()
        $Script:install_command | Out-File $Intunewin_Command_File -Force -NoNewline
        $Form_Intunewin.close()
    })

    $Form_Intunewin.Add_Closing({
        $Script:install_command = $install_command_intunewin.Text.ToString()
        $Script:install_command | Out-File $Intunewin_Command_File -Force -NoNewline
    })

    $Form_Intunewin.ShowDialog() | Out-Null

    $Sandbox_Root_Path = "C:\Run_in_Sandbox"
    $Intunewin_Installer = "$Sandbox_Root_Path\IntuneWin_Install.ps1"
    return ($PSRun_File + " " + "$Intunewin_Installer")
}

# Show VBS Params dialog and return the startup command
function Show-VBSParamsDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Run_in_Sandbox_Folder,
        [Parameter(Mandatory=$true)]
        [string]$Full_Startup_Path_UnQuoted
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Config.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($XamlPath)
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlLoader)
    $Form_VBS = [Windows.Markup.XamlReader]::Load($Reader)

    $parameters_to_add = $Form_VBS.findname("parameters_to_add")
    $add_parameters = $Form_VBS.findname("add_parameters")

    $Script:DialogStartupCommand = ""
    
    $add_parameters.add_click({
        $Script:Paramaters = $parameters_to_add.Text.ToString()
        $Script:DialogStartupCommand = "wscript.exe `"$Full_Startup_Path_UnQuoted`" " + $Script:Paramaters
        $Form_VBS.close()
    })

    $Form_VBS.Add_Closing({
        $Script:Paramaters = $parameters_to_add.Text.ToString()
        $Script:DialogStartupCommand = "wscript.exe `"$Full_Startup_Path_UnQuoted`" " + $Script:Paramaters
    })

    $Form_VBS.ShowDialog() | Out-Null
    
    return $Script:DialogStartupCommand
}

Export-ModuleMember -Function @(
    'Show-EXEDialog',
    'Show-MSIDialog',
    'Show-PS1ParamsDialog',
    'Show-IntunewinDialog',
    'Show-VBSParamsDialog'
)
