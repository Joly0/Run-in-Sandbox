# Run in Sandbox: a quick way to run/extract files in Windows Sandbox from a right-click
###### *[View the full blog post here](https://www.systanddeploy.com/2023/06/runinsandbox-quick-way-to-runextract.html)*

#### Original Author & creator: Damien VAN ROBAEYS
#### Rewritten and maintained now by: Joly0

This allows you to do the below things in Windows Sandbox **just from a right-click** by adding context menus:
- Run PS1 as user or system in Sandbox
- Run CMD, VBS, EXE, MSI in Sandbox
- Run Intunewin file
- Open URL or HTML file in Sandbox
- Open PDF file in Sandbox
- Extract ZIP file directly in Sandbox
- Extract 7z file directly in Sandbox (uses host-installed 7-Zip or downloads latest version)
- Extract ISO directly in Sandbox (uses host-installed 7-Zip or downloads latest version)
- Share a specific folder in Sandbox
- Run multiple app´s/scripts in the same Sandbox session

![alt text](https://github.com/damienvanrobaeys/Run-in-Sandbox/blob/master/ps1_system.gif)

**Note that this project has been build on personal time, it's not a professional project. Use it at your own risk, and please read How to install it before running it.**
<br/>
<br/>
<br/>
<br/>

# How to install it ?
All the steps need to be executed from the Host, not inside the Sandbox

### <mark>__Method 1 - PowerShell (Recommended)__</mark>
-   Right-click on the Windows start menu and select PowerShell or Terminal (Not CMD), preferably as admin.
-   Copy and paste the code below and press enter:

#### **Install from master branch (stable):**
```powershell
irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1 | iex
```

#### **Install from a specific branch (e.g., dev):**
```powershell
iex "& { $(irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1) } -Branch dev"
```

Replace `dev` with your desired branch name (e.g., `beta`, `test`, etc.)

#### **Additional installer parameters:**
The installer supports the following optional parameters:

| Parameter | Description |
|-----------|-------------|
| `-Branch <name>` | Install from a specific branch (default: `master`, or your currently installed branch for updates) |
| `-DeepClean` | Performs a deep-clean of legacy registry entries before installation. Removes old context menu entries (takes 5-10 minutes) |
| `-NoCheckpoint` | Skips creation of a system restore point during installation |

**Examples:**
```powershell
# Install from dev branch with deep-clean
iex "& { $(irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1) } -Branch dev -DeepClean"

# Install without creating a restore point
iex "& { $(irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1) } -NoCheckpoint"

# Combine multiple parameters
iex "& { $(irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1) } -Branch dev -DeepClean -NoCheckpoint"
```

-   You will see the process being started. You will probably be asked to grant admin rights if not started as admin.
-   That's all.

Note - On older Windows builds you may need to run the below command first: \
```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
```
<br/>
<br/>

### <mark>__Method 2 - Traditional__</mark>
This method allows you to use the parameters "-NoCheckpoint" to skip creation of a restore point and "-NoSilent" to give a bit more output
- Download the ZIP Run-in-Sandbox project (this is the main prerequiste)
- Extract the ZIP
- The Run-in-Sandbox-master **should contain** at least Add_Structure.ps1  and a Sources folder
- Please **do not download only** Add_Structure.ps1
- The Sources folder **should contain** folder Run_in_Sandbox containing 58 files
- Once you have downloaded the folder structure, **check if files have not be blocked after download**
- Do a right-click on Add_Structure.ps1 and check if needed check Unblocked
- Run Add_Structure.ps1 **with admin rights**
<br/>

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=Joly0/Run-in-Sandbox&type=Date)](https://www.star-history.com/#Joly0/Run-in-Sandbox&Date)
