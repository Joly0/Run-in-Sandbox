<#
.SYNOPSIS
    Dialog functions module for Run-in-Sandbox

.DESCRIPTION
    This module provides dialog functionality for the Run-in-Sandbox application.
    It handles displaying various dialog boxes for user interaction.
#>

function Show-EXEDialog {
    param(
        [string]$Run_in_Sandbox_Folder,
        [string]$Full_Startup_Path_Quoted
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_EXE.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    Function LoadXml ($global:file2) {
        $XamlLoader = (New-Object System.Xml.XmlDocument)
        $XamlLoader.Load($file2)
        return $XamlLoader
    }

    $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_EXE.xaml")
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
    $Form_EXE = [Windows.Markup.XamlReader]::Load($Reader)
    $EXE_Command_File = "$Run_in_Sandbox_Folder\EXE_Command_File.txt"

    $switches_for_exe = $Form_EXE.findname("switches_for_exe")
    $add_switches = $Form_EXE.findname("add_switches")

    $add_switches.Add_Click({
        $Script:Switches_EXE = $switches_for_exe.Text.ToString()
        $Script:Startup_Command = $Full_Startup_Path_Quoted + " " + $Switches_EXE
        $Startup_Command | Out-File $EXE_Command_File -Force -NoNewline
        $Form_EXE.close()
    })

    $Form_EXE.Add_Closing({
        $Script:Switches_EXE = $switches_for_exe.Text.ToString()
        $Script:Startup_Command = $Full_Startup_Path_Quoted + " " + $Switches_EXE
        $Startup_Command | Out-File $EXE_Command_File -Force -NoNewline
    })

    $Form_EXE.ShowDialog() | Out-Null

    $EXE_Installer = "$Sandbox_Root_Path\EXE_Install.ps1"
    $Script:Startup_Command = $PSRun_File + " " + "$EXE_Installer"
    
    $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
    New-WSB -Command_to_Run $Startup_Command
}

function Show-MSIDialog {
    param(
        [string]$Run_in_Sandbox_Folder,
        [string]$Full_Startup_Path_UnQuoted
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Config.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    Function LoadXml ($global:file2) {
        $XamlLoader = (New-Object System.Xml.XmlDocument)
        $XamlLoader.Load($file2)
        return $XamlLoader
    }

    $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Config.xaml")
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
    $Form_MSI = [Windows.Markup.XamlReader]::Load($Reader)

    $switches_for_exe = $Form_MSI.findname("switches_for_exe")
    $add_switches = $Form_MSI.findname("add_switches")

    $add_switches.Add_Click({
        $Script:Switches_MSI = $switches_for_exe.Text.ToString()
        $Script:Startup_Command = "msiexec /i `"$Full_Startup_Path_UnQuoted`" " + $Switches_MSI
        $Form_MSI.close()
    })

    $Form_MSI.Add_Closing({
        $Script:Switches_MSI = $switches_for_exe.Text.ToString()
        $Script:Startup_Command = "msiexec /i `"$Full_Startup_Path_UnQuoted`" " + $Switches_MSI
    })

    $Form_MSI.ShowDialog() | Out-Null
    
    $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
    New-WSB -Command_to_Run $Startup_Command
}

function Show-ParamsDialog {
    param(
        [string]$Run_in_Sandbox_Folder,
        [string]$Full_Startup_Path_UnQuoted,
        [string]$PSRun_File
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    Function LoadXml ($global:file2) {
        $XamlLoader = (New-Object System.Xml.XmlDocument)
        $XamlLoader.Load($file2)
        return $XamlLoader
    }

    $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml")
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
    $Form_PS1 = [Windows.Markup.XamlReader]::Load($Reader)

    $Form_PS1.Add_Loaded({
        # Set up event handlers here if needed
    })
    
    $Form_PS1.Add_Closing({
        # Handle closing event
    })

    $Form_PS1.ShowDialog() | Out-Null
    
    $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
    New-WSB -Command_to_Run $Startup_Command
}

function Show-IntunewinDialog {
    param(
        [string]$Run_in_Sandbox_Folder
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Intunewin.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    Function LoadXml ($global:file2) {
        $XamlLoader = (New-Object System.Xml.XmlDocument)
        $XamlLoader.Load($file2)
        return $XamlLoader
    }

    $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Intunewin.xaml")
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
    $Form_PS1 = [Windows.Markup.XamlReader]::Load($Reader)

    $install_command_intunewin = $Form_PS1.findname("install_command_intunewin")
    $add_install_command = $Form_PS1.findname("add_install_command")

    $add_install_command.add_click({
        $Script:install_command = $install_command_intunewin.Text.ToString()
        $install_command | Out-File $Intunewin_Command_File
        $Form_PS1.close()
    })

    $Form_PS1.Add_Closing({
        $Script:install_command = $install_command_intunewin.Text.ToString()
        $install_command | Out-File $Intunewin_Command_File -Force -NoNewline
        $Form_PS1.close()
    })

    $Form_PS1.ShowDialog() | Out-Null

    $Intunewin_Installer = "$Sandbox_Root_Path\IntuneWin_Install.ps1"
    $Script:Startup_Command = $PSRun_File + " " + "$Intunewin_Installer"
    
    $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
    New-WSB -Command_to_Run $Startup_Command
}

function Show-VBSParamsDialog {
    param(
        [string]$Run_in_Sandbox_Folder,
        [string]$Full_Startup_Path_UnQuoted
    )
    
    $XamlPath = "$Run_in_Sandbox_Folder\RunInSandbox_Config.xaml"
    if (-not (Test-Path $XamlPath)) {
        throw "XAML file not found: $XamlPath"
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
    [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
    
    Function LoadXml ($global:file2) {
        $XamlLoader = (New-Object System.Xml.XmlDocument)
        $XamlLoader.Load($file2)
        return $XamlLoader
    }

    $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Config.xaml")
    $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
    $Form_VBS = [Windows.Markup.XamlReader]::Load($Reader)

    $parameters_to_add = $Form_VBS.findname("parameters_to_add")
    $add_parameters = $Form_VBS.findname("add_parameters")

    $add_parameters.add_click({
        $Script:Paramaters = $parameters_to_add.Text.ToString()
        $Script:Startup_Command = "wscript.exe $Full_Startup_Path_UnQuoted $Paramaters"
        $Form_VBS.close()
    })

    $Form_VBS.Add_Closing({
        $Script:Paramaters = $parameters_to_add.Text.ToString()
        $Script:Startup_Command = "wscript.exe $Full_Startup_Path_UnQuoted $Paramaters"
    })

    $Form_VBS.ShowDialog() | Out-Null
    
    $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
    New-WSB -Command_to_Run $Startup_Command
}

Export-ModuleMember -Function @(
    'Show-EXEDialog',
    'Show-MSIDialog',
    'Show-ParamsDialog',
    'Show-IntunewinDialog',
    'Show-VBSParamsDialog'
)