# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## 2026-10-09
### Fixed
- Fixed `Remove_Structure.ps1` not removing the `.img` context menu added in 2026-08-07
- Fixed installation aborting when `HKCU\Software\Classes\.msix` exists without an `OpenWithProgids` subkey, and MSIX entries being written to a wrong key when more than one ProgID is registered
- Fixed "Extract ISO file in Sandbox" pointing the sandbox to the host path of the cached 7-Zip installer (only used when 7-Zip is not installed on the host)
- Fixed the "Failed to download 7-Zip installer" message never being shown when the download failed and no cached installer existed
- Fixed the sandbox not starting at all when no classic `notepad.exe.mui` can be found (the Notepad payload is optional now)
- Fixed application bundles (`.sdbapp`) not finding their files: folders are now mapped to `C:\SBDApp\<folder name>` as expected by `AppBundle_Install.ps1`; `.ps1`/`.vbs` paths with spaces and a missing `Silent_Switch` element are handled
- Fixed files in folders with `[` or `]` in their name being started from a wrong path in the sandbox
- Fixed the installer always using `master` for updates instead of the installed branch when `-Branch` is not given
- Fixed deep-clean updates deleting custom startup scripts and the just created backup folder
- Fixed update creating nested empty folders (e.g. `Modules\Shared\Shared`)
- Fixed deep-clean possibly scanning the wrong user's `HKCU_Classes` hive
- Fixed the Windows Sandbox feature check never reaching its fallback when `Get-WindowsOptionalFeature` fails, and the disk space check assuming drive `C:`
- Fixed `RunInSandbox_Config.ps1` (`Add_Structure.ps1 -NoSilent`) failing to load its assemblies and XAML when started from another folder
- Fixed `Remove_Structure.ps1` reporting success although the installation folder could not be removed
- Fixed the registry backup (`Registry_Backup` folder) never being created: `Export-RegConfig` returned before exporting anything. Each modified key is now exported once, before the first change, keys containing `\` no longer point into a folder that does not exist, HKCR and HKU backups of the same key no longer overwrite each other, and an existing backup is kept on updates
- Fixed the HTML/URL entry: the path was passed with its double quotes inside single quotes (`-LiteralPath '"C:\..."'`), so `Invoke-Item` never found the file
- Fixed the HTML entry showing twice for Edge/Chrome users: it is only added to `SystemFileAssociations` now (which applies to every default browser), and `.htm` and `.url` are covered as well. The browser ProgID entries of older versions are still removed on uninstall
- Fixed the ZIP entry missing when another application (7-Zip, PeaZip, WinZip, ...) is the default for .zip files: it is also added to the ProgID of that application
- Fixed "Run EXE with switches" breaking on file names with spaces (`EXE_Install.ps1`)
- Fixed `IntuneWin_Install.ps1` failing when the dialog wrote a trailing new line
- Fixed application bundles (`.sdbapp`) with another file name than `App_Bundle.sdbapp` not being found by `AppBundle_Install.ps1`
- Fixed the launch failing for standard users when `_orchestrator.ps1` (installed by the elevated installer) could not be overwritten: it is only written when missing or outdated now, and the existing one is used if that fails. A failing write of `OriginalCommand.txt` reports the path instead of starting a sandbox that does nothing
- Notepad payload: a missing `notepad.exe.mui` no longer skips the whole payload (only the localized strings are missing then), the 0 byte app execution alias in WindowsApps is not copied any more, and `01-Copy-Notepad.ps1` no longer registers "Edit with Notepad" when no notepad.exe was staged
- Fixed downloads of branches containing `/` (the extracted folder is read from the archive now)
- Fixed the installer "falling back" to `CommonFunctions.ps1` when a module could not be downloaded, which lacks most installer functions (`Resolve-Branch`, `Invoke-AsAdmin`, `Write-Info`, ...) and then failed with unrelated errors. It now names the module that failed and exits

### Added
- The sandbox uses **Notepad++** from the host when it is installed: its folder is mounted read only to `C:\Program Files\Notepad++` (like the host installation of 7-Zip), and inside the sandbox "Edit with Notepad++", "Open Notepad++" on the folder background and the .txt association are registered. Only without Notepad++ the classic Notepad is staged
- `RunInSandbox.ps1` writes its log to `%TEMP%\RunInSandbox.log` and shows a message box on errors, unknown `-Type` values and a missing .wsb file, instead of failing silently because it is started with `-WindowStyle Hidden`
- Added `-RepoOwner` and `-RepoName` parameters to `Install_Run-in-Sandbox.ps1` to install from a fork. The repository is stored in `version.json`, so updates stay on it; default is still `Joly0/Run-in-Sandbox`


## 2026-08-07
### Added
- Added .img file to "Run X in Sandbox"-options


## 2026-05-13
### Fixed
- Fixed `Add_Structure.ps1` not adding the ps1 context menu when the UserChoice ProgID has no `Shell` subkey in HKCR (e.g. after removing PowerShell ISE the gate `Test-Path "$HKCR_UserChoice_Key\Shell"` returned `$False` and the menu was silently skipped). The cascade is now written to `HKCU\Software\Classes\<ProgID>\Shell` so the menu shows up via the merged HKCR view even without a machine-wide entry, and the whole block is wrapped in try/catch with a null guard on the UserChoice ProgID. Fixes [#20]https://github.com/Joly0/Run-in-Sandbox/issues/20
- Mirrored the above ps1 UserChoice fix in `Remove_Structure.ps1` so the uninstall removes the entries from both the new `HKCU_Classes\<ProgID>\Shell` location and the legacy HKCR one for older installs
- Fixed Intunewin context menu failing with `Cannot dot-source this command because it was defined in a different language mode` when AppLocker enforces ConstrainedLanguage on scripts under `C:\ProgramData\Run_in_Sandbox\`. The registry verb command now invokes `RunInSandbox.ps1` via `-Command "& '...'"` instead of `-File`, which internally dot-sources and triggers the language mode mismatch
- Fixed Intunewin install dialog silently auto-closing on some locked-down Windows 11 builds (focus-stealing prevention / ASR rules can tear down the WPF/MahApps window before the `+` click handler runs). Added a `Windows.Forms` InputBox fallback that takes over when the WPF dialog produces no install command
- Fixed `IntuneWin_Install.ps1` silently exiting inside the sandbox when no install command was captured. Replaced the silent `EXIT` with visible MessageBoxes so the reason is shown on the sandbox desktop


## 2026-04-30a
### Fixed
- Fixed regression bug with previous bug fix, should finally solve [#18]https://github.com/Joly0/Run-in-Sandbox/issues/18

## 2026-04-30
### Fixed
- Fixed bugged installation for ps1 context menu´s when different admin account is needed for UAC (ps1 context menu was written to admin registry). Fixes [#18]https://github.com/Joly0/Run-in-Sandbox/issues/18


## 2026-03-30
### Fixed
- Fixed logging file path being affected by Windows 8.3 short paths when username has space should fix [#17]https://github.com/Joly0/Run-in-Sandbox/issues/17


## 2026-03-04
### Fixed
- Fixed `_orchestrator.ps1` access denied error when starting a sandbox after a fresh install
- Fixed cleanup of Intunewin and EXE temp files not being removed when sandbox is closed
- Fixed installation not stopping on errors (replaced `EXIT` with `throw` in validation functions for proper error propagation)
### Changed
- Moved runtime temp files (`Intunewin_Folder.txt`, `Intunewin_Install_Command.txt`, `EXE_Command_File.txt`) into the `temp/` subfolder
- Installer now sets targeted `Modify` permissions for BUILTIN\Users on `temp/`, `startup-scripts/`, and `Sandbox_Config.xml` instead of `FullControl` on the entire install folder
### Added
- `Set-UserWritePermissions` function in `Environment.psm1` for reusable, language-independent permission management
- Pre-install prerequisite checks for RAM (≥4 GB) and disk space (≥1 GB free)
- Option to automatically enable Windows Sandbox feature if not installed, with user confirmation and manual reboot
- `$ErrorActionPreference = 'Stop'` in `Add_Structure.ps1` to halt on any error


## 2026-02-02
### Fixed
- Fixed dialog forVBSParams and MSI files


## 2026-02-02
### Changed
- Properly implement hiding and showing shell (CMD and Powershell) windows inside the Sandbox
### Added
- Added debug output for the different shells when running with showing shell windows (`AppBundle_Install.ps1`,`IntuneWin_Install.ps1`)
- Added additional winforms window in error case with additional error information
### Fixed
- Fixed various bugs in the `RunInSandbox.ps1` script


## 2025-12-19
### Changed
- **Code Restructuring:** Refactored codebase to use PowerShell modules for better maintainability
  - Created `Modules/Shared/` for common utilities (Logging, Config, Environment, Version)
  - Created `Modules/Runtime/` for sandbox execution (WSB, SevenZip, StartupScripts, Dialogs, UI)
  - Created `Modules/Installer/` for installation logic (Registry, Validation, Core)
- **Advanced Installer:** Completely rewritten `Install_Run-in-Sandbox.ps1` with new features:
  - `-Branch` parameter for installing from different branches (master, dev)
  - `-DeepClean` parameter for thorough cleanup of legacy registry entries
  - `-NoCheckpoint` parameter to skip system restore point creation
  - Dynamic module loading from GitHub during installation
  - Version display and reinstall prompt for existing installations
  - Automatic backup creation before updates
  - Fallback to CommonFunctions.ps1 for backward compatibility
  - Updated documentation in the readme to reflect changes and new parameters
### Added
- `version.json` for tracking installed version and branch
- `Modules/Installer/*` scripts with advanced installation functions
- `Modules/Shared/*` with shared scripts
- `Modules/Runtime/*` with runtime specific scripts
- `04-Install_VSRedist.ps1` script to startup script to install Visual Studio Redistributables on statup (often required for software installations)
### Removed
- Dependency on `CommonFunctions.ps1` for all scripts (for backwards compatiblity will remain in the code base for now)


## 2025-12-08
### Fixed
- Fixed running "folder on" and "folder in" https://github.com/Joly0/Run-in-Sandbox/issues/12


## 2025-10-07
### Added
- Added HTML file association when other browsers (instead of chrome or edge) are default
### Changed
- Improved Orchestrator script
- Improved notepad startup script
### Fixed
- Added missing startup-scripts folder


## 2025-10-06
### Added
- Added startup functionality, making it possible to run multiple startup Scripts in the Sandbox. Improves (and kinda Fixes) [#11]https://github.com/Joly0/Run-in-Sandbox/issues/11
- Added some useful startup script (Notepad, Context Menu customization, fix for slow MSI installation); Thanks to ThioJoe https://github.com/ThioJoe/Windows-Sandbox-Tools
- Added code to unblock files in the host folder that is mapped to the sandbox. Fixes [#10]https://github.com/Joly0/Run-in-Sandbox/issues/10
### Fixed
- Fixed running multiple Apps through SDBApp
### Changed
- Changed some startup behaviour and adjusted showing some cmd/powershell windows and hiding some


## 2025-08-20
### Security
- **Enhanced 7-Zip Integration:** Removed bundled 7-Zip executables for improved security
- **On-Demand 7-Zip:** Now uses host-installed 7-Zip when available, or downloads latest version from GitHub releases
- **Smart Caching:** Automatic download and caching of 7-Zip installers with 7-day refresh cycle
- **Offline Support:** Works offline using cached installers when network is unavailable
- **Always Current:** Ensures latest 7-Zip version is used, eliminating security vulnerabilities from outdated bundled files
### Changed
- Modified `RunInSandbox.ps1` to detect and mount host 7-Zip installation
- Updated installation process to cache latest 7-Zip installer during setup
- Enhanced `New-WSB` function to support additional mapped folders
- Reduced project size by ~2MB by removing bundled 7-Zip files
### Added
- `Find-Host7Zip` function for detecting system 7-Zip installations
- `Get-Latest7ZipDownloadUrl` function for GitHub API integration
- `Update-7ZipCache` function for smart installer management
- `Ensure-7ZipCache` function for offline-capable cache validation
- Version tracking and age-based cache refresh mechanism


## 2025-05-07
### Fixed
- Fixed running batch files in Sandbox
- Fixed label for "Run BAT in Sandbox"


## 2025-01-06
### Fixed
- Fixed [#7]https://github.com/Joly0/Run-in-Sandbox/issues/7 and [#58]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/58
- Fixed indendation for wsb file
- Fixed [#8]https://github.com/Joly0/Run-in-Sandbox/issues/8


## 2024-11-8
### Changed
- Slightly adjusted RunInSandbox installer script
- Improved readme for RunInSandbox installation


## 2024-10-15
### Fixed
- Fixed [#56]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/56
### Changed
- Formatting improvements for RunInSandbox.ps1 script


## 2024-08-26
### Fixed
- Fixed Subcommands Entries for Powershell and VBS


## 2024-08-20
### Added
- Added common functions script
- Added deep-clean option for uninstalling Run-in-Sandbox
### Changed
- Rewritten install and uninstall script and exported common functions to separate file
### Fixed
- Slightly adjusted intunewin and sdbapp scripts


## 2024-05-22
### Added
- Added easy install Script
### Changed
- Improved console output and make it easier readable
- Improved readme and install steps


## 2024-05-14
### Added
-  Added some better error handling and checking for needed features
### Fixed
- Probably fixed [#4]https://github.com/Joly0/Run-in-Sandbox/issues/4
### Changed
- Improved the way, exe files are handled inside the sandbox


## 2023-07-14
### Fixed
- Finally fixed running intunewin with serviceUI and psexec
- Fixed [#40]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/40
- Fixed [#41]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/41
### Changed
- Changed formatting to OTBS using "Invoke-Formatter" cmdlet in "Script-Analyzer" module (On-going discussion [#44]https://github.com/damienvanrobaeys/Run-in-Sandbox/discussions/44) and applied some powershell best-pratices


## 2023-05-01
### Added
- Reimplemented running Intunewin as System using psexec (serviceui will stay)
### Fixed
- Fixed [#18]https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/18


## 2023-05-03
### Added
- Added option to run .intunewin via sdbapp
### Changed
- Changed Intunewin_Content_File and Intunewin_Command_File to be parameters for IntuneWin_Install.ps1


## 2023-03-29
### Added
- Added context menu entry for opening PDF in Sandbox
### Changed
- Completely rewrote alot of code in Add_Structure.ps1


## 2023-03-22
### Added
- Added context menu entry for running CMD/BAT in Sandbox


## 2023-03-21
### Changed
- Readded 7z part and adjusted 7z reg key path


## 2023-03-20
### Changed
- Completly refactored RunInSandbox.ps1 to use switch instead of ifelse and rearranged alot of code
### Fixed
- Fixed some issues with loading iso´s, exe´s and zip´s
### Removed
- Removed 7z part of RunInSandbox.ps1 because non-functional


## 2023-03-07
### Added
- Added ServiceUI
### Changed
- Replaced PSexec with ServiceUI for intunewin sandbox
### Removed
- Removed PSexec in favor of ServiceUI


## 2023-03-06
### Added
- Added option to Sandbox_Config.xml to cleanup leftover .wsb file afterwards (default is true)
### Changed
- .wsb is not executed by the "Start-Process"-cmdlet with -wait parameter


## 2023-03-03
### Added
- Added -noprofile to powershell commands to improve performance
### Changed
- Applied formatting of scripts and applied best practices
### Fixed
- Fixed .ps1 conext menu


## 2021-11-16
### Added
- Add a context menu for running PS1 as system in Sandbox
- Add a context menu for running MSIX in Sandbox
- Add a context menu for running PPKG in Sandbox
- Add a context menu for opening URL in Sandbox
- Add a context menu for extracting ISO in Sandbox
- Add a context menu for extracting 7z file in Sandbox
### Fixed
- Fix a bug where context menu for PS1 does not appear on Windows 11 


## 2021-09-21
### Added
- Add a context menu for reg file, to run them in Sandbox
- Add ability to run multiple apps in the same Sandbox session


## 2021-08-03
### Added
- Add a context menu for intunewin file, to run them in Sandbox
- Add ability to choose which content menu to add


## 2021-07-27
### Changed
- Change the default path where WSB are saved after running Sandbox: now in %temp%


## 2021-07-21
### Changed
- Updated the GUI when running EXE or MSI for more understanding
- Updated the GUI when running PS1 for more understanding


## 2021-07-16
### Added
- The Add_Structure.ps1 will now create a restore point
- It will then check if Sources folder exists


## 2020-06-24
### Removed
- Temporarily removed the main file [#9](https://github.com/damienvanrobaeys/Run-in-Sandbox/issues/9)
### Changed
- Fixed detail language setting being French


## 2020-06-02 
### Added
 - Add new WSB config options for Windows 10 2004. These new settings can be managed in the **Sources\Run_in_Sandbox\Sandbox_Config.xml**
 - New options: AudioInput, VideoInput, ProtectedClient, PrinterRedirection, ClipboardRedirection, MemoryInMB


## 2020-05-19
### Added
- Added French, Italian, Spanish, English, and German languages for context menus. To configure language, edit **Main_Language** in **Sources\Run_in_Sandbox\Sandbox_Config.xml**
