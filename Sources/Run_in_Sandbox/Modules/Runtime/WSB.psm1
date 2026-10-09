<#
.SYNOPSIS
    Windows Sandbox (WSB) management module for Run-in-Sandbox

.DESCRIPTION
    This module provides Windows Sandbox functionality for the Run-in-Sandbox application.
    It handles creation, configuration, and management of Windows Sandbox environments.
#>

function New-WSB {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [String]$Command_to_Run,
        
        [Array]$AdditionalMappedFolders = @(),
        
        [string]$FileName = "Sandbox",
        
        [string]$DirectoryName = "",
        
        [string]$ScriptPath = "",
        
        [string]$Type = ""
    )
    
    $Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
    $xml = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
    $my_xml = [xml](Get-Content $xml)
    $Sandbox_VGpu = $my_xml.Configuration.VGpu
    $Sandbox_Networking = $my_xml.Configuration.Networking
    $Sandbox_ReadOnlyAccess = $my_xml.Configuration.ReadOnlyAccess
    $Sandbox_WSB_Location = $my_xml.Configuration.WSB_Location
    $Sandbox_AudioInput = $my_xml.Configuration.AudioInput
    $Sandbox_VideoInput = $my_xml.Configuration.VideoInput
    $Sandbox_ProtectedClient = $my_xml.Configuration.ProtectedClient
    $Sandbox_PrinterRedirection = $my_xml.Configuration.PrinterRedirection
    $Sandbox_ClipboardRedirection = $my_xml.Configuration.ClipboardRedirection
    $Sandbox_MemoryInMB = $my_xml.Configuration.MemoryInMB
    
    # Give the sandbox an editor. Notepad++ from the host is preferred and only mounted
    # read only, like the host installation of 7-Zip (01-Copy-Notepad.ps1 registers it).
    # Only without Notepad++ the classic Notepad is staged. Both are optional and must not
    # prevent the sandbox from starting
    $NotepadPlusPlus_Path = Find-HostNotepadPlusPlus
    if ($NotepadPlusPlus_Path) {
        $AdditionalMappedFolders += @{
            HostFolder = Split-Path $NotepadPlusPlus_Path -Parent
            SandboxFolder = "C:\Program Files\Notepad++"
            ReadOnly = "true"
        }
        Write-LogMessage -Message_Type "INFO" -Message "Using the Notepad++ installation of the host: $NotepadPlusPlus_Path"
    } else {
        try {
            Add-NotepadToSandbox -EnforceEnUsFallback
        } catch {
            Write-LogMessage -Message_Type "WARNING" -Message "No Notepad++ found and Notepad could not be prepared either: $($_.Exception.Message)"
        }
    }
    
    if ($Sandbox_WSB_Location -eq "Default") {
        $Sandbox_File_Path = "$env:temp\$FileName.wsb"
    } else {
        $Sandbox_File_Path = "$Sandbox_WSB_Location\$FileName.wsb"
    }

    if (Test-Path $Sandbox_File_Path) {
        Remove-Item $Sandbox_File_Path
    }
    
    New-Item $Sandbox_File_Path -type file -Force | Out-Null
    Add-Content -LiteralPath $Sandbox_File_Path -Value "<Configuration>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <VGpu>$Sandbox_VGpu</VGpu>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <Networking>$Sandbox_Networking</Networking>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <AudioInput>$Sandbox_AudioInput</AudioInput>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <VideoInput>$Sandbox_VideoInput</VideoInput>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <ProtectedClient>$Sandbox_ProtectedClient</ProtectedClient>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <PrinterRedirection>$Sandbox_PrinterRedirection</PrinterRedirection>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    <ClipboardRedirection>$Sandbox_ClipboardRedirection</ClipboardRedirection>"
    if ( -not [string]::IsNullOrEmpty($Sandbox_MemoryInMB) ) {
        Add-Content -LiteralPath $Sandbox_File_Path -Value "    <MemoryInMB>$Sandbox_MemoryInMB</MemoryInMB>"
    }

    Add-Content $Sandbox_File_Path "    <MappedFolders>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>C:\ProgramData\Run_in_Sandbox</HostFolder>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\Run_in_Sandbox</SandboxFolder>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
    Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"

    if ($Type -eq "SDBApp" -and $ScriptPath) {
        $SDB_Full_Path = $ScriptPath
        # AppBundle_Install.ps1 reads the bundle as "App_Bundle.sdbapp", whatever the
        # file the user picked is called
        Copy-Item $ScriptPath "$Run_in_Sandbox_Folder\App_Bundle.sdbapp" -Force
        $Get_Apps_to_install = [xml](Get-Content $SDB_Full_Path)
        $Apps_to_install_path = $Get_Apps_to_install.Applications.Application.Path | Select-Object -Unique

        ForEach ($App_Path in $Apps_to_install_path) {
            Get-ChildItem -LiteralPath $App_Path -Recurse | Unblock-File
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$App_Path</HostFolder>"
            # AppBundle_Install.ps1 expects each app under C:\SBDApp\<last folder of Path>
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\SBDApp\$($App_Path.Split('\')[-1])</SandboxFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
        }
    } elseif ($DirectoryName) {
        Get-ChildItem -LiteralPath $DirectoryName -Recurse | Unblock-File
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$DirectoryName</HostFolder>"
        if ($Type -eq "IntuneWin") { Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\IntuneWin</SandboxFolder>" }
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    
    # Add any additional mapped folders
    foreach ($MappedFolder in $AdditionalMappedFolders) {
        Get-ChildItem -LiteralPath $($MappedFolder.HostFolder) -Recurse | Unblock-File
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$($MappedFolder.HostFolder)</HostFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>$($MappedFolder.SandboxFolder)</SandboxFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$($MappedFolder.ReadOnly)</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    </MappedFolders>"
    
    if ( -not [string]::IsNullOrEmpty($Command_to_Run) ) {
        Add-Content -LiteralPath $Sandbox_File_Path  -Value "    <LogonCommand>"
        Add-Content -LiteralPath $Sandbox_File_Path  -Value "        <Command>$Command_to_Run</Command>"
        Add-Content -LiteralPath $Sandbox_File_Path  -Value "    </LogonCommand>"
    }

    Add-Content -LiteralPath $Sandbox_File_Path  -Value "</Configuration>"
}

# Finds a Notepad++ installation on the host system
function Find-HostNotepadPlusPlus {
    $CommonPaths = @(
        "${env:ProgramFiles}\Notepad++\notepad++.exe",
        "${env:ProgramFiles(x86)}\Notepad++\notepad++.exe"
    )
    foreach ($Path in $CommonPaths) {
        if ( (-not [string]::IsNullOrEmpty($Path)) -and (Test-Path -LiteralPath $Path) ) {
            return $Path
        }
    }

    # Registry and uninstall entry, these also find installations on another drive
    $InstallFolders = @(
        (Get-ItemProperty -Path "HKLM:\SOFTWARE\Notepad++" -ErrorAction SilentlyContinue)."(default)"
        (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Notepad++" -ErrorAction SilentlyContinue)."(default)"
        (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Notepad++" -ErrorAction SilentlyContinue).InstallLocation
        (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Notepad++" -ErrorAction SilentlyContinue).InstallLocation
    )
    foreach ($Install_Folder in $InstallFolders) {
        if ( (-not [string]::IsNullOrEmpty($Install_Folder)) -and (Test-Path -LiteralPath "$Install_Folder\notepad++.exe") ) {
            return "$Install_Folder\notepad++.exe"
        }
    }

    # PATH, this also covers portable installations added to PATH
    $NotepadPlusPlusInPath = Get-Command "notepad++.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($NotepadPlusPlusInPath) {
        return $NotepadPlusPlusInPath.Source
    }

    return $null
}

function Remove-Leftovers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RemovalPath
    )
    if (Test-Path $RemovalPath) {
        Remove-Item -LiteralPath $RemovalPath -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function @(
    'New-WSB',
    'Find-HostNotepadPlusPlus',
    'Remove-Leftovers'
)
