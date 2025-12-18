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
    
    # Prepare Notepad payload
    Add-NotepadToSandbox -EnforceEnUsFallback
    
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
        Copy-Item $ScriptPath $Run_in_Sandbox_Folder -Force
        $Get_Apps_to_install = [xml](Get-Content $SDB_Full_Path)
        $Apps_to_install_path = $Get_Apps_to_install.Applications.Application.Path | Select-Object -Unique

        ForEach ($App_Path in $Apps_to_install_path) {
            Get-ChildItem -LiteralPath $App_Path -Recurse | Unblock-File
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$App_Path</HostFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\SBDApp</SandboxFolder>"
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
    'Remove-Leftovers'
)
