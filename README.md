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
##### __`irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/master/Install_Run-in-Sandbox.ps1 | iex`__

#### **Install from a specific branch (e.g., dev):**
##### __`iex "& { $(irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/dev/Install_Run-in-Sandbox.ps1) } -Branch dev"`__

Replace both `dev` instances with your desired branch name (e.g., `beta`, `test`, etc.)

-   You will see the process being started. You will probably be asked to grant admin rights if not started as admin.
-   That's all.

Note - On older Windows builds you may need to run the below command first: \
__`[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12`__
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

## Reliable branch selection when running via iex

Passing parameters through the classic iex pattern can fail to bind -Branch to the downloaded script, which causes the installer to fall back to version.json. Use one of the following robust patterns to ensure the branch is honored.

- Option A — Robust parameter binding with ScriptBlock (recommended):
  - Replace dev with your target branch.
  - Works in Windows PowerShell 5.1+ and PowerShell 7+.
  ```
  iex '(&{ $sb = [ScriptBlock]::Create((irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/dev/Install_Run-in-Sandbox.ps1)); & $sb -Branch "dev" })'
  ```

- Option B — Environment variable override (simple and reliable):
  - Set an environment variable that the installer reads before version.json. Replace dev with your target branch.
  ```
  $env:RIS_BRANCH = 'dev'
  iex "& { $(irm https://raw.githubusercontent.com/Joly0/Run-in-Sandbox/dev/Install_Run-in-Sandbox.ps1) }"
  ```

Behavior and precedence
- The installer in [Install_Run-in-Sandbox.ps1](Install_Run-in-Sandbox.ps1) resolves branch using this precedence:
  1) -Branch parameter (if bound)
  2) Environment variable RIS_BRANCH or RUN_IN_SANDBOX_BRANCH
  3) Invocation parsing fallbacks (may not work in all iex styles)
  4) Existing installation&#39;s version.json
  5) Default master

Notes
- If you use the basic iex pattern with a subshell, the -Branch argument may not bind to the script&#39;s param() and can be ignored. Use Option A or B above to guarantee the desired branch is used.
- All URLs in your command should match the desired branch (don&#39;t mix dev URL with master branch value, or vice versa).
