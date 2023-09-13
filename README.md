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

## Background / Why?
WinGet have a few limitations in terms of automation and is not integrated with common endpoints management products.  
System Context is not possible by using "winget" as the .exe location must be found and this location is not static due to versioning in the directory name.

## Requirements
Newer Windows OS build that includes the WinGet natively in the OS   
Windows 10 20H2 or newer should be enough

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

## Usage (InTune)
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
