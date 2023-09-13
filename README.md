# WinGet-Wrapper  
PowerShell Scripts to install/manage applications using WinGet - Common use case include InTune, and other endpoint management products.
Detection script automatically check local installed version against latest winget available version or a defined fixed target version.
Easily import finished applications including WinGet metadata using WinGet-WrapperImportFromCSV.ps1 script

* Dynamically finds the WinGet directory to be used under System Context.
* Kill selected process before WinGet command
* Allows running pre and post script before installation
* Detection script that dynamically finds latest package available trough WinGet
* Requirement script to allow creating packages for update purposes only
* Logs to $env:ProgramData\WinGet-WrapperLogs (Usually C:\ProgramData\WinGet-WrapperLogs)
* Dynamically detect if running in user or system context
* Performs automatic cleanup of log files older than 60 days.
* Directly import WinGet packages to InTune including WinGet package metadata

## Background / Why?
WinGet have a few limitations in terms of automation and is not integrated with common endpoints management products.  
System Context is not possible by using "winget" as the .exe location must be found and this location is not static due to versioning in the directory name.

## Requirements
Newer Windows OS build that includes the WinGet natively in the OS   
Windows 10 20H2 or newer should be enough
Powershell 5.1
WinGet-WrapperImportFromCSV.ps1 automatically installs required module "IntuneWin32App" (github.com/MSEndpointMgr/IntuneWin32App)

## WinGet-Wrapper.ps1
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/4220b44b-7f96-4fb1-84ec-ce416f6f622c)

#### Usage:
>Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "PackageName for log file" -StopProcess "kill process using Stop-Process (do not add .exe)" -PreScript "somefile.ps1" -PostScript "somefile.ps1" -ArgumentList "Arguments Passed to WinGet.exe"

## WinGet-WrapperDetection.ps1
Matches locally installed version with newest available version using WinGet or specified version using $TargetVersion
Can be setup to accept newer installed version locally $AcceptNewerVersion
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/631d6001-b813-4b79-a12f-3c1e06cb3aec)

## WinGet-WrapperRequirements.ps1
Checks if application is detected locally. If not detected will not attempt update/install
To be used when only wanting to update if application is already installed. (Update Only)

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b5af0ddd-6700-46cf-8907-33dbd0f8e930)

Outputs either "Installed" or "Not Installed"

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b8cd24fd-da34-4e1c-aeb2-0627717e1244)

## WinGet-WrapperImportFromCSV.ps1
Imports packages from WinGet to InTune (incuding available WinGet package metadata)
Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/c626ed5b-80eb-4d56-8476-605349356ffa)

## Usage Import from CSV (InTune)
Open the sample CSV file WinGet-WrapperImportFromCSV.csv and add any WinGet Package IDs to import (Case Sensitive)

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
#### Usage:
>WinGet-WrapperImportFromCSV.ps1 -TenantID company.onmicrosoft.com -csvFile WinGet-WrapperImportFromCSV.csv -SkipConfirmation

## Usage Manual Import (InTune)
**Application Installation**

In InTune create an Windows app (Win32) and upload WinGet-Wrapper.InTuneWin as the package file.  
>**Install:** Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"

>**Uninstall:** Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"

Change the $id variable to match the package id in the detection script and upload it  ($id = "VideoLAN.VLC")  
  *If specific version is required change the $TargetVersion (Ex. $TargetVersion = "1.0.0.0")*  
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/2aea611c-7733-4f93-9cbe-a44b4f66333d)

**Application Update Only**

For creating application that will only update/install if application is already installed
Perform the same steps as in "Application Installation".
Setup Requirement rule script with return string value of "Installed"

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b2bdb617-c74a-4902-9c2c-b8defe1adc70)

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b8cd24fd-da34-4e1c-aeb2-0627717e1244)
