# Define global variables (preserving original lines 1-19)
$TEMP_Folder = $env:temp
$Log_File = "$TEMP_Folder\RunInSandbox_Install.log"
$Run_in_Sandbox_Folder = "$env:ProgramData\Run_in_Sandbox"
$XML_Config = "$Run_in_Sandbox_Folder\Sandbox_Config.xml"
$Windows_Version = (Get-CimInstance -class Win32_OperatingSystem).Caption
$currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
if (Test-Path -LiteralPath "Registry::HKEY_USERS\$currentSid\Volatile Environment" -ErrorAction SilentlyContinue) {
    $Current_User_SID = $currentSid
    $HKCU = "Registry::HKEY_USERS\$Current_User_SID"
    $HKCU_Classes = "Registry::HKEY_USERS\${Current_User_SID}_Classes"
} else {
    $HKCU = 'HKCU:'
    $HKCU_Classes = 'HKCU:\Software\Classes'
}
$Sandbox_Icon = "$env:ProgramData\Run_in_Sandbox\sandbox.ico"
# $Current_Folder is not defined in this context, so we'll use $ScriptDir instead
# This ensures the Sources variable works correctly
$Sources = "$ScriptDir\Sources\*"
$Exported_Keys = @()

# Load System.Windows.Forms assembly (preserving line 20)
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

# Get the directory where this script is located
# This allows the script to find the modules regardless of where it's called from
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import all of the necessary modules
# Shared modules
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Shared\Logging.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Shared\Config.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Shared\Environment.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Shared\Version.psm1" -Force

# Installer modules
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Installer\Registry.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Installer\Core.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Installer\Validation.psm1" -Force

# Runtime modules
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Runtime\SevenZip.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Runtime\Update.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Runtime\StartupScripts.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Runtime\WSB.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Runtime\Dialogs.psm1" -Force
Import-Module "$ScriptDir\Sources\Run_in_Sandbox\Modules\Runtime\UI.psm1" -Force

# Note: Functions from imported modules are automatically available when this script is dot-sourced
# Global variables are already defined above and will be available to scripts that dot-source this file

# This shim ensures backward compatibility with existing scripts that do `. .\CommonFunctions.ps1`
# The actual function implementations are now in the modular structure
