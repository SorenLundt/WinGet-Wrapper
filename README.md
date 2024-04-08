# WinGet-Wrapper  
PowerShell Scripts to deploy and bulk import WinGet packages to InTune including metadata.<br>
Automatically detect latest version using dynamic detection script. <br>
Detection script checks local installed version against latest winget available version or a defined fixed target version.<br>
Bulk import WinGet packages to InTune including WinGet package metadata using WinGet-WrapperImportGUI.exe <br>
<br>
* Dynamically finds the WinGet directory to be used under System Context.<br>
* Kill selected process before WinGet command<br>
* Allows running pre and post script before installation<br>
* Detection script that dynamically finds latest package available trough WinGet<br>
* Requirement script to allow creating packages for update purposes only<br>
* Logs to $env:ProgramData\WinGet-WrapperLogs (Usually C:\ProgramData\WinGet-WrapperLogs)<br>
* Dynamically detect if running in user or system context<br>
* Performs automatic cleanup of log files older than 60 days.<br>
* Directly import and deploy WinGet packages to InTune including WinGet package metadata<br>

## Background / Why?
WinGet have a few limitations in terms of automation and is not integrated with common endpoints management products.  <br>
System Context is not possible by using "winget" as the .exe location must be found and this location is not static due to versioning in the directory name.<br>

## Requirements
Windows 10 20H2 or newer<br>
Powershell 5.1<br>
Client language must be en-US, as Winget-Wrapper parses only English output. <br>
Module "IntuneWin32App" and "Microsoft.Graph.Intune" needed for import to InTune <br>

## WinGet-WrapperImportGUI.exe
WinGet-WrapperImportGUI is a graphical interface designed to streamline the import of WinGet packages into InTune. <br>
This tool complements WinGet-Wrapper, providing an intuitive way to upload WinGet packages to InTune, along with their metadata. <br>

#### Features:
- **Search and Select:** Seamlessly search for WinGet packages and select the ones you need.
- **InTune Integration:** Import selected WinGet packages directly into InTune for deployment.
- **CSV Support:** Export and import packages using CSV files, facilitating batch operations.<br>

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/1b43c387-3adf-4eb7-90c3-99dcf07b9871)



#### Usage:
>**Open the GUI:** Run WinGet-WrapperImportGUI.exe to open the GUI<br>
>**Search Packages:** Enter your search query and click "Search" to find WinGet packages.<br>
>**Select Packages:** Select from search results, then click the center arrow to move them to the import list.<br>
>**Import to InTune:** Enter your Tenant ID and click "Import to InTune" to import selected packages.<br>
>**Additional Actions:** Use buttons for exporting CSV, deleting, or importing from CSV.<br>

If you get errors when using the ImportGUI, you may need to unblock winget-wrapper files from file properties.<br>
WinGet-WrapperImportFromCSV.ps1, IntuneWinAppUtil.exe, WinGet-WrapperImportGUI.ps1, more..
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/9dc24e0b-966b-4aee-bfbc-e28235d1bcfb)


## WinGet-Wrapper.ps1
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/4220b44b-7f96-4fb1-84ec-ce416f6f622c)

#### Usage:
>Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "PackageName for log file" -StopProcess "kill process using Stop-Process (do not add .exe)" -PreScript "somefile.ps1" -PostScript "somefile.ps1" -ArgumentList "Arguments Passed to WinGet.exe"

## WinGet-WrapperDetection.ps1
Matches locally installed version with newest available version using WinGet or specified version using $TargetVersion<br>
Can be setup to accept newer installed version locally $AcceptNewerVersion<br><br>
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/631d6001-b813-4b79-a12f-3c1e06cb3aec)

## WinGet-WrapperRequirements.ps1
Checks if application is detected locally. If not detected will not attempt update/install<br>
To be used when only wanting to update if application is already installed. (Update Only)<br><br>
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b5af0ddd-6700-46cf-8907-33dbd0f8e930)

Outputs either "Installed" or "Not Installed"<br><br>
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b8cd24fd-da34-4e1c-aeb2-0627717e1244)

## WinGet-WrapperImportFromCSV.ps1
Imports packages from WinGet to InTune (incuding available WinGet package metadata)<br>
Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm<br>
Create deployment using csv columns: InstallIntent, Notification, GroupID<br><br>
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/dde433b9-81e1-484b-8ee1-71ac02d68441)
<br>
## Usage: Import from CSV (InTune)
Open the sample CSV file WinGet-WrapperImportFromCSV.csv and add any WinGet Package IDs to import (Case Sensitive)<br>
#### Usage:
>WinGet-WrapperImportFromCSV.ps1 -TenantID company.onmicrosoft.com -csvFile WinGet-WrapperImportFromCSV.csv -SkipConfirmation
#### Process:
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/297ddb07-eeac-41c7-a9ec-9656727984f6)
<br>
#### Results:
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/ca57c9d4-0ec7-4514-8694-7160f6356b5e)


#### Columns:
* PackageID = Exact PackageID (Required)<br>
* Context = Which context the Win32App is run under (Machine or User) (Required)<br>
* AcceptNewerVersion = Allows newer installed version locally than specified (Set to 0 or 1)(Required)<br>
* UpdateOnly = Update package only. Application will only update if application is already installed (Set to 0 or 1)(Required)<br>
* TargetVersion = Specfic version of the application. If not set, the package will always be the latest version <br>
* StopProcessInstall = Kill a specific process (Stop-process) before installation (.exe should not be defined)<br>
* StopProcessUninstall = Kill a specific process (Stop-process) before uninstallation (.exe should not be defined)<br>
* PreScriptInstall = Run powershell script before installation<br>
* PostScript = Run powershell script after installation<br>
* PreScriptUninstall = Run powershell script before uninstallation<br>
* PostScriptUninstall = Run powershell script after uninstallation<br>
* CustomArgumentListInstall = Arguments passsed to WinGet (default: install --exact --id PackageID --silent --accept-package-agreements --accept-source-agreements --scope Context<br>
* CustomArgumentListUninstall = Arguments passsed to WinGet (default: uninstall --exact --id PackageID --silent --accept-source-agreements --scope Context<br>
* InstallIntent = Available or Required deployment <br>
* Notification = Notification level on deployment - Valid values: showAll, showReboot, hideAll<br>
* GroupID = InTune GroupID to deploy package to<br>


## Usage: Manual Import (InTune)
**Application Installation**

In InTune create an Windows app (Win32) and upload WinGet-Wrapper.InTuneWin as the package file.  <br>
>**Install:** Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"

>**Uninstall:** Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"

Change the $id variable to match the package id in the detection script and upload it  ($id = "VideoLAN.VLC")  <br>
  *If specific version is required change the $TargetVersion (Ex. $TargetVersion = "1.0.0.0")*  <br><br>
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/2aea611c-7733-4f93-9cbe-a44b4f66333d)

**Application Update Only**

For creating application that will only update/install if application is already installed<br>
Perform the same steps as in "Application Installation".<br>
Setup Requirement rule script with return string value of "Installed"<br><br>
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b2bdb617-c74a-4902-9c2c-b8defe1adc70)

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b8cd24fd-da34-4e1c-aeb2-0627717e1244)


## Disclaimer
This software is provided "AS IS" with no warranties. Use at your own risk.
