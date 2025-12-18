param (
    [Parameter(Mandatory=$true)] [String]$Type,
    [Parameter(Mandatory=$true)] [String]$ScriptPath
)

#Start-Transcript -Path $(Join-Path -Path $([System.Environment]::GetEnvironmentVariables('Machine').TEMP) -ChildPath "RunInSandbox.log")

$special_char_array = 'é', 'è', 'à', 'â', 'ê', 'û', 'î', 'ä', 'ë', 'ü', 'ï', 'ö', 'ù', 'ò', '~', '!', '@', '#', '$', '%', '^', '&', '+', '=', '}', '{', '|', '<', '>', ';'
foreach ($char in $special_char_array) {
    if ($ScriptPath -like "*$char*") {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $message = "There is a special character in the path of the file (`'" + $char + "`').`nWindows Sandbox does not support this!"
        [System.Windows.Forms.MessageBox]::Show($message, "Issue with your file", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        EXIT
    }
}

$ScriptPath = $ScriptPath.replace('"', '')
$ScriptPath = $ScriptPath.Trim();
$ScriptPath = [WildcardPattern]::Escape($ScriptPath)

if ( ($Type -eq "Folder_Inside") -or ($Type -eq "Folder_On") ) {
    $DirectoryName = (Get-Item $ScriptPath).fullname
} else {
    $FolderPath = Split-Path -LiteralPath (Split-Path -LiteralPath "$ScriptPath" -Parent) -Leaf
    $DirectoryName = (Get-Item $ScriptPath).DirectoryName
    $FileName = (Get-Item $ScriptPath).BaseName
    $Full_FileName = (Get-Item $ScriptPath).Name
}

$Sandbox_Desktop_Path = "C:\Users\WDAGUtilityAccount\Desktop"
$Sandbox_Shared_Path = "$Sandbox_Desktop_Path\$FolderPath"

$Sandbox_Root_Path = "C:\Run_in_Sandbox"
$Full_Startup_Path = "$Sandbox_Shared_Path\$Full_FileName"
$Full_Startup_Path_Quoted = """$Full_Startup_Path"""

$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"

# Import required modules
Import-Module "$Run_in_Sandbox_Folder\Modules\Shared\Logging.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Shared\Config.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\SevenZip.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\WSB.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\StartupScripts.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\UI.psm1" -Force
Import-Module "$Run_in_Sandbox_Folder\Modules\Runtime\Dialogs.psm1" -Force

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
$WSB_Cleanup = $my_xml.Configuration.WSB_Cleanup
$Hide_Powershell = $my_xml.Configuration.Hide_Powershell

[System.Collections.ArrayList]$PowershellParameters = @(
    '-sta'
    '-WindowStyle'
    'Hidden'
    '-NoProfile'
    '-ExecutionPolicy'
    'Unrestricted'
)

if ($Hide_Powershell -eq "False") {
    $PowershellParameters[[array]::IndexOf($PowershellParameters, "Hidden")] = "Normal"
}

$PSRun_File = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $PowershellParameters -File"
$PSRun_Command = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $PowershellParameters -Command"

if ($Sandbox_WSB_Location -eq "Default") {
    $Sandbox_File_Path = "$env:temp\$FileName.wsb"
} else {
    $Sandbox_File_Path = "$Sandbox_WSB_Location\$FileName.wsb"
}

if (Test-Path -LiteralPath $Sandbox_File_Path) {
    Remove-Item $Sandbox_File_Path
}

switch ($Type) {
    "7Z" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()

        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path -LiteralPath $Host7ZipPath -Parent

            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }

            $Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File"

            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Initialize-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }

            $CachedInstaller = "$Sandbox_Root_Path\temp\7zSetup.msi"

            # Install 7-Zip in sandbox then extract
            $Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File' -Wait`""
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer: $CachedInstaller"
        }

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "CMD" {
        $Startup_Command = $PSRun_Command + " " + "Start-Process $Full_Startup_Path_Quoted"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "EXE" {
        $Startup_Command = Show-EXEDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_Quoted $Full_Startup_Path_Quoted -PSRun_File $PSRun_File

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "Folder_On" {
        New-WSB -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "Folder_Inside" {
        New-WSB -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "HTML" {
        $Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -LiteralPath `'$Full_Startup_Path_Quoted`'`""

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "URL" {
        $Startup_Command = $PSRun_Command + " " + "Start-Process $Sandbox_Root_Path"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "Intunewin" {
        $Startup_Command = Show-IntunewinDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -FileName $FileName -PSRun_File $PSRun_File

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "ISO" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()

        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path -LiteralPath $Host7ZipPath -Parent

            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }

            $Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO"

            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation for ISO: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Initialize-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }

            $CachedInstaller = "$Run_in_Sandbox_Folder\temp\7zSetup.msi"

            # Install 7-Zip in sandbox then extract ISO
            $Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO' -Wait`""
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer for ISO: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "MSI" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")
        $Startup_Command = Show-MSIDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_UnQuoted $Full_Startup_Path_UnQuoted

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "MSIX" {
        $Startup_Command = $PSRun_Command + " " + "Add-AppPackage -LiteralPath $Full_Startup_Path_Quoted"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PDF" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -LiteralPath `'$Full_Startup_Path_Quoted`'`""

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PPKG" {
        $Startup_Command = $PSRun_Command + " " + "Install-ProvisioningPackage $Full_Startup_Path_Quoted -forceinstall -quietinstall"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PS1Basic" {
        $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PS1System" {
        $Startup_Command = "$Sandbox_Root_Path\PsExec.exe \\localhost -nobanner -accepteula -s Powershell -ExecutionPolicy Bypass -File $Full_Startup_Path_Quoted"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "PS1Params" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")
        $Startup_Command = Show-PS1ParamsDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_UnQuoted $Full_Startup_Path_UnQuoted -PSRun_File $PSRun_File

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "REG" {
        $Startup_Command = "REG IMPORT $Full_Startup_Path_Quoted"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "SDBApp" {
        $AppBundle_Installer = "$Sandbox_Root_Path\AppBundle_Install.ps1"
        $Startup_Command = $PSRun_File + " " + "$AppBundle_Installer"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "VBSBasic" {
        $Startup_Command = "wscript.exe $Full_Startup_Path_Quoted"

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "VBSParams" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $Startup_Command = Show-VBSParamsDialog -Run_in_Sandbox_Folder $Run_in_Sandbox_Folder -Full_Startup_Path_UnQuoted $Full_Startup_Path_UnQuoted

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
    "ZIP" {
        $Startup_Command = $PSRun_Command + " " + "`"Expand-Archive -LiteralPath '$Full_Startup_Path' -DestinationPath '$Sandbox_Desktop_Path\ZIP_extracted'`""

        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -FileName $FileName -DirectoryName $DirectoryName -ScriptPath $ScriptPath -Type $Type
    }
}

Start-Process -FilePath $Sandbox_File_Path -Wait
do {
    Start-Sleep -Seconds 1
} while (Get-Process -Name "WindowsSandboxServer" -ErrorAction SilentlyContinue)

if ($WSB_Cleanup -eq $True) {
    Remove-Leftovers -RemovalPath $Sandbox_File_Path
    Remove-Leftovers -RemovalPath $Intunewin_Command_File
    Remove-Leftovers -RemovalPath $Intunewin_Content_File
    Remove-Leftovers -RemovalPath $EXE_Command_File
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\App_Bundle.sdbapp"
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\NotepadPayload"
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\startup-scripts\OriginalCommand.txt"
}
