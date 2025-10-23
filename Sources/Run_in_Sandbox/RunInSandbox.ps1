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
    $FolderPath = Split-Path (Split-Path "$ScriptPath" -Parent) -Leaf
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

# Load common functions

# Start asynchronous update check (non-blocking)
Start-Job -ScriptBlock {
    param($Run_in_Sandbox_Folder)
    . "$Run_in_Sandbox_Folder\CommonFunctions.ps1"
    Start-UpdateCheck
} -ArgumentList $Run_in_Sandbox_Folder | Out-Null
. "$Run_in_Sandbox_Folder\CommonFunctions.ps1"

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

if (Test-Path $Sandbox_File_Path) {
    Remove-Item $Sandbox_File_Path
}

function Enable-StartupScripts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OriginalCommand,  # whatever you would have passed as -Command_to_Run
        [string]$StartupScriptFolderName = "startup-scripts"
    )

    $StartupScriptsFolder = Join-Path $Run_in_Sandbox_Folder $StartupScriptFolderName
    New-Item -ItemType Directory -Path $StartupScriptsFolder -Force | Out-Null
    
    if ($OriginalCommand -ne "") {
       # Write the original command into a file
        $origCmdFile = Join-Path $StartupScriptsFolder "OriginalCommand.txt"
        Set-Content -LiteralPath $origCmdFile -Value $OriginalCommand -Encoding UTF8 -Force 
    }

    # Orchestrator that runs NN-*.ps1 in lexicographic order, then runs the original command
    $orchestrator = @'
param(
    [string]$ScriptsPath = "C:\Run_in_Sandbox\startup-scripts",
    # Can include this switch when running from the .wsb file to indicate it's the first launch of the sandbox
    # Useful if re-running this script within the sandbox as a test, but don't want certain parts to run again
    [switch]$launchingSandbox
)

# ------ Check that we're running in the Windows Sandbox ------
# This script is intended to be run from within the Windows Sandbox. We'll do a rudamentary check for if the current user is named "WDAGUtilityAccount"
if ($env:USERNAME -ne "WDAGUtilityAccount") {
    Write-host "`n`nERROR: This script is intended to be run from WITHIN the Windows Sandbox.`nIt appears you are running this from outside the sandbox.`n" -ForegroundColor Red
    Write-host "`nPress Enter to exit." -ForegroundColor Yellow
    Read-Host
    exit
}

Write-Host "[Orchestrator] Scripts path: $ScriptsPath"

# 1) Run ordered startup scripts: 00-*, 01-* ... 99-*
$pattern = '^\d{2}-.+\.ps1$'
$items = Get-ChildItem -LiteralPath $ScriptsPath -Filter *.ps1 -File -ErrorAction SilentlyContinue |
         Where-Object { $_.Name -match $pattern } |
         Sort-Object Name

foreach ($i in $items) {
    Write-Host "[Orchestrator] Running: $($i.Name)"
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $i.FullName
        $rc = $LASTEXITCODE
        if ($rc -ne $null -and $rc -ne 0) {
            Write-Warning "[Orchestrator] Script $($i.Name) returned exit code $rc"
        }
    } catch {
        Write-Warning "[Orchestrator] Script $($i.Name) threw: $($_.Exception.Message)"
    }
}

# Restart Explorer so changes take effect
Write-Host "[Orchestrator] Restarting Explorer so changes take effect"
Get-Process explorer | Stop-Process -Force

# 2) Read and run the original command last
$origFile = Join-Path $ScriptsPath "OriginalCommand.txt"
if (Test-Path -LiteralPath $origFile) {
    $orig = Get-Content -LiteralPath $origFile -Raw
    Write-Host "[Orchestrator] Running original command..."
    # Run through cmd to support both cmd and PowerShell-style lines
    Start-process -Filepath "C:\Windows\SysWOW64\cmd.exe" -ArgumentList @('/c', '"' + $orig + '"') -WindowStyle Hidden
} else {
    Write-Warning "[Orchestrator] OriginalCommand.txt not found; nothing to run."
}
'@

    $orchestratorPath = Join-Path $StartupScriptsFolder "_orchestrator.ps1"
    Set-Content -LiteralPath $orchestratorPath -Value $orchestrator -Encoding UTF8 -Force

    # Return the single Sandbox command that runs the orchestrator
    "C:\Run_in_Sandbox\ServiceUI.exe -Process:explorer.exe C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Normal -NoProfile -ExecutionPolicy Bypass -NoExit -File `"$Sandbox_Root_Path\$StartupScriptFolderName\_orchestrator.ps1`""
}

function Add-NotepadToSandbox {
    [CmdletBinding()]
    param(
        [string]$HostPayloadRoot = "C:\ProgramData\Run_in_Sandbox\NotepadPayload",
        [switch]$EnforceEnUsFallback # if set, will try en-US when preferred language is missing
    )

    # Resolve a single notepad.exe (prefer System32)
    $exeCandidates = Get-Command notepad.exe -ErrorAction Stop | Select-Object -ExpandProperty Source
    $exePath = ($exeCandidates | Where-Object { $_ -match '\\Windows\\System32\\' } | Select-Object -First 1)
    if (-not $exePath) { $exePath = $exeCandidates | Select-Object -First 1 }

    $exeDir  = Split-Path $exePath -Parent
    $exeName = Split-Path $exePath -Leaf

    # Build candidate language list
    $candidates = @()
    try { $candidates += (Get-UICulture).Name } catch {}
    try { $candidates += (Get-WinSystemLocale).Name } catch {}
    try {
        $candidates += Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\MUI\UILanguages" |
                       Select-Object -ExpandProperty PSChildName
    } catch {}
    $candidates = $candidates | Where-Object { $_ } | Select-Object -Unique

    # Probe possible MUI locations
    $dirs = @(
        $exeDir,
        (Join-Path $env:WINDIR 'System32'),
        (Join-Path $env:WINDIR 'SysWOW64'),
        $env:WINDIR
    ) | Select-Object -Unique

    $muiPath = $null; $resolvedLang = $null
    foreach ($lang in $candidates) {
        foreach ($dir in $dirs) {
            $p = Join-Path (Join-Path $dir $lang) "$exeName.mui"
            if (Test-Path -LiteralPath $p) { $muiPath = $p; $resolvedLang = $lang; break }
        }
        if ($muiPath) { break }
    }

    if (-not $muiPath -and $EnforceEnUsFallback) {
        foreach ($dir in $dirs) {
            $fallback = Join-Path (Join-Path $dir 'en-US') "$exeName.mui"
            if (Test-Path -LiteralPath $fallback) { $muiPath = $fallback; $resolvedLang = 'en-US'; break }
        }
    }

    if (-not $muiPath) {
        throw "Could not locate notepad.exe.mui for $exePath. On some systems Notepad is a Store app without a classic MUI."
    }

    # Stage payload on host: System32\notepad.exe and System32\<lang>\notepad.exe.mui
    $sys32Out = Join-Path $HostPayloadRoot "System32"
    $langOut  = Join-Path $sys32Out $resolvedLang
    New-Item -ItemType Directory -Path $langOut -Force | Out-Null
    Copy-Item -LiteralPath $exePath -Destination (Join-Path $sys32Out $exeName) -Force
    Copy-Item -LiteralPath $muiPath -Destination (Join-Path $langOut "$exeName.mui") -Force
}

function New-WSB {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [String]$Command_to_Run,
        [Array]$AdditionalMappedFolders = @()
    )
    
    # Prepare Notepad payload
    Add-NotepadToSandbox -EnforceEnUsFallback
    
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

    if ($Type -eq "SDBApp") {
        $SDB_Full_Path = $ScriptPath
        Copy-Item $ScriptPath $Run_in_Sandbox_Folder -Force
        $Get_Apps_to_install = [xml](Get-Content $SDB_Full_Path)
        $Apps_to_install_path = $Get_Apps_to_install.Applications.Application.Path | Select-Object -Unique

        ForEach ($App_Path in $Apps_to_install_path) {
            Get-ChildItem -Path $App_Path -Recurse | Unblock-File
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$App_Path</HostFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\SBDApp</SandboxFolder>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
            Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
        }
    } else {
        Get-ChildItem -Path $DirectoryName -Recurse | Unblock-File
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$DirectoryName</HostFolder>"
        if ($Type -eq "IntuneWin") { Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>C:\IntuneWin</SandboxFolder>" }
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$Sandbox_ReadOnlyAccess</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    
    # Add any additional mapped folders
    foreach ($MappedFolder in $AdditionalMappedFolders) {
        Get-ChildItem -Path $($MappedFolder.HostFolder) -Recurse | Unblock-File
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        <MappedFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <HostFolder>$($MappedFolder.HostFolder)</HostFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <SandboxFolder>$($MappedFolder.SandboxFolder)</SandboxFolder>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "            <ReadOnly>$($MappedFolder.ReadOnly)</ReadOnly>"
        Add-Content -LiteralPath $Sandbox_File_Path -Value "        </MappedFolder>"
    }
    Add-Content -LiteralPath $Sandbox_File_Path -Value "    </MappedFolders>"
    
    Add-Content -Path $Sandbox_File_Path  -Value "    <LogonCommand>"
    Add-Content -Path $Sandbox_File_Path  -Value "        <Command>$Command_to_Run</Command>"
    Add-Content -Path $Sandbox_File_Path  -Value "    </LogonCommand>"
    Add-Content -Path $Sandbox_File_Path  -Value "</Configuration>"
}

switch ($Type) {
    "7Z" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()
        
        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path $Host7ZipPath -Parent
            
            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }
            
            $Script:Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File"
            
            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Ensure-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }
            
            $CachedInstaller = "$Sandbox_Root_Path\temp\7zSetup.msi"
            
            # Install 7-Zip in sandbox then extract
            $Script:Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_File' -Wait`""
            
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders
    }
    "CMD" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "EXE" {
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
    "Folder_On" {
        $Startup_Command = Enable-StartupScripts
        New-WSB -Command_to_Run $Startup_Command
    }
    "Folder_Inside" {
        $Startup_Command = Enable-StartupScripts
        New-WSB -Command_to_Run $Startup_Command
    }
    "HTML" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "URL" {
        $Script:Startup_Command = $PSRun_Command + " " + "Start-Process $Sandbox_Root_Path"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "Intunewin" {
        $Intunewin_Folder = "C:\IntuneWin\$FileName.intunewin"
        $Intunewin_Content_File = "$Run_in_Sandbox_Folder\Intunewin_Folder.txt"
        $Intunewin_Command_File = "$Run_in_Sandbox_Folder\Intunewin_Install_Command.txt"
        $Intunewin_Folder | Out-File $Intunewin_Content_File -Force -NoNewline

        #$Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')  | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        function LoadXml ($global:file1) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file1)
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
    "ISO" {
        # Try to find 7-Zip on host system first
        $Host7ZipPath = Find-Host7Zip
        $AdditionalFolders = @()
        
        if ($Host7ZipPath) {
            # Mount the host 7-Zip installation into sandbox
            $Host7ZipFolder = Split-Path $Host7ZipPath -Parent
            
            $AdditionalFolders += @{
                HostFolder = $Host7ZipFolder
                SandboxFolder = "C:\Program Files\7-Zip"
                ReadOnly = "true"
            }
            
            $Script:Startup_Command = "`"C:\Program Files\7-Zip\7z.exe`" x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO"
            
            Write-LogMessage -Message_Type "INFO" -Message "Using host 7-Zip installation for ISO: $Host7ZipPath"
        }
        else {
            # No host installation found, ensure we have a cached installer
            if (-not (Ensure-7ZipCache)) {
                [System.Windows.Forms.MessageBox]::Show("Failed to download 7-Zip installer and no cached version available.`nPlease check your internet connection.")
                EXIT
            }
            
            $CachedInstaller = "$Run_in_Sandbox_Folder\temp\7zSetup.msi"
            
            # Install 7-Zip in sandbox then extract ISO
            $Script:Startup_Command = "$PSRun_Command `"Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i \`"$CachedInstaller\`" /quiet' -Wait; Start-Process -FilePath 'C:\Program Files\7-Zip\7z.exe' -ArgumentList 'x $Full_Startup_Path_Quoted -y -oC:\Users\WDAGUtilityAccount\Desktop\Extracted_ISO' -Wait`""
            
            Write-LogMessage -Message_Type "INFO" -Message "Using cached 7-Zip installer for ISO: $CachedInstaller"
        }
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command -AdditionalMappedFolders $AdditionalFolders
    }
    "MSI" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')              | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll")      | Out-Null
        function LoadXml ($global:file2) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file2)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_EXE.xaml")
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
    "MSIX" {
        $Script:Startup_Command = $PSRun_Command + " " + "Add-AppPackage -Path $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PDF" {
        $Full_Startup_Path_Quoted = $Full_Startup_Path_Quoted.Replace('"', '')
        $Script:Startup_Command = $PSRun_Command + " " + "`"Invoke-Item -Path `'$Full_Startup_Path_Quoted`'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PPKG" {
        $Script:Startup_Command = $PSRun_Command + " " + "Install-ProvisioningPackage $Full_Startup_Path_Quoted -forceinstall -quietinstall"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1Basic" {
        $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1System" {
        $Script:Startup_Command = "$Sandbox_Root_Path\PsExec.exe \\localhost -nobanner -accepteula -s Powershell -ExecutionPolicy Bypass -File $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "PS1Params" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', "")

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')  | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        function LoadXml ($global:file1) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file1)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml")
        $Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
        $Form_PS1 = [Windows.Markup.XamlReader]::Load($Reader)

        $parameters_to_add = $Form_PS1.findname("parameters_to_add")
        $add_parameters = $Form_PS1.findname("add_parameters")

        $add_parameters.add_click({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_UnQuoted" + " " + "$Paramaters"
                $Form_PS1.close()
            })

        $Form_PS1.Add_Closing({
                $Script:Paramaters = $parameters_to_add.Text.ToString()
                $Script:Startup_Command = $PSRun_File + " " + "$Full_Startup_Path_UnQuoted" + " " + "$Paramaters"
            })

        $Form_PS1.ShowDialog() | Out-Null
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "REG" {
        $Script:Startup_Command = "REG IMPORT $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "SDBApp" {
        $AppBundle_Installer = "$Sandbox_Root_Path\AppBundle_Install.ps1"
        $Script:Startup_Command = $PSRun_File + " " + "$AppBundle_Installer"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "VBSBasic" {
        $Script:Startup_Command = "wscript.exe $Full_Startup_Path_Quoted"
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
    "VBSParams" {
        $Full_Startup_Path_UnQuoted = $Full_Startup_Path_Quoted.Replace('"', '')

        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.dll") | Out-Null
        [System.Reflection.Assembly]::LoadFrom("$Run_in_Sandbox_Folder\assembly\MahApps.Metro.IconPacks.dll") | Out-Null
        function LoadXml ($Script:file1) {
            $XamlLoader = (New-Object System.Xml.XmlDocument)
            $XamlLoader.Load($file1)
            return $XamlLoader
        }

        $XamlMainWindow = LoadXml("$Run_in_Sandbox_Folder\RunInSandbox_Params.xaml")
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
    "ZIP" {
        $Script:Startup_Command = $PSRun_Command + " " + "`"Expand-Archive -LiteralPath '$Full_Startup_Path' -DestinationPath '$Sandbox_Desktop_Path\ZIP_extracted'`""
        
        $Startup_Command = Enable-StartupScripts -OriginalCommand $Startup_Command
        New-WSB -Command_to_Run $Startup_Command
    }
}

Start-Process -FilePath $Sandbox_File_Path -Wait
do {
    Start-Sleep -Seconds 1
} while (Get-Process -Name "WindowsSandboxServer" -ErrorAction SilentlyContinue)

if ($WSB_Cleanup -eq $True) {
    function Remove-Leftovers {
        param(
            [Parameter(Mandatory=$true)]
            [string]$RemovalPath = ""
        )
        if (Test-Path $RemovalPath) {
            Remove-Item -LiteralPath $RemovalPath -Force -Recurse -ErrorAction SilentlyContinue
        }
    }
    Remove-Leftovers -RemovalPath $Sandbox_File_Path
    Remove-Leftovers -RemovalPath $Intunewin_Command_File
    Remove-Leftovers -RemovalPath $Intunewin_Content_File
    Remove-Leftovers -RemovalPath $EXE_Command_File
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\App_Bundle.sdbapp"
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\NotepadPayload"
    Remove-Leftovers -RemovalPath "$Run_in_Sandbox_Folder\startup-scripts\OriginalCommand.txt"
}

# Check for pending update installation (runs regardless of cleanup setting)
$UpdateDecision = Get-UpdateDecision
if ($UpdateDecision -and $UpdateDecision.action -eq "update") {
    Write-Host "Installing update to version $($UpdateDecision.latestVersion)..." -ForegroundColor Green
    Write-Host "Starting update process in new window..." -ForegroundColor Cyan
    
    # Download and execute Install_Run-in-Sandbox.ps1 as a separate process
    try {
        $VersionInfo = Get-VersionInfo
        $Branch = $VersionInfo.CurrentBranch
        $InstallScriptUrl = "https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/$Branch/Install_Run-in-Sandbox.ps1"
        
        # Download to temp file
        $TempInstaller = "$env:TEMP\Run-in-Sandbox-Updater.ps1"
        Invoke-WebRequest -Uri $InstallScriptUrl -OutFile $TempInstaller -UseBasicParsing
        
        # Launch installer in new elevated process (don't wait)
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$TempInstaller`" -AutoUpdate -NoCheckpoint" -Verb RunAs
        
        Write-Host "Update process started. This window will now close." -ForegroundColor Green
        Write-Host "The update will complete in the new window." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        
        # Exit immediately to release file locks
        exit 0
    } catch {
        Write-Host "Failed to start update process: $_" -ForegroundColor Red
        Write-Host "You can manually update by running: powershell -c `"irm https://run-in-sandbox.com/install.ps1 | iex`"" -ForegroundColor Yellow
    }
}