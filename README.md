# WinGet-Wrapper  
A PowerShell script to install/manage applications using WinGet - Common use case include InTune, and other endpoint management products.´ 
Detection script available for fixed version and for dynamic version (automatically matches the installed version with version available with WinGet)

* Can Dynamically find the WinGet directory to be used under System Context.
* Installation under either User and System context is supported
* Kill selected process before WinGet command
* Detection script that dynamically finds latest package available trough WinGet
* Requirement script to create packages that performs update only if application is detected
* Logs to $env:ProgramData\WinGet-WrapperLogs (Usually C:\ProgramData\WinGet-WrapperLogs) 
* Detection script performs automatic cleanup of log files older than 60 days.

## Background / Why?
WinGet have a few limitations in terms of automation and is not integrated with common endpoints management products.  
System Context is not possible by using "winget" as the .exe location must be found and this location is not static due to versioning in the directory name.

## Requirements
Newer Windows OS build that includes the WinGet natively in the OS   
Windows 10 20H2 or newer should be enough

## WinGet-Wrapper.ps1
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/fa0a68a2-b790-489f-8995-fd26d8031f55)
#### Usage:
>Powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -Context "User or System Context" -StopProcess "kill process using Stop-Process (do not add .exe)" -ArgumentList "Arguments Passed to WinGet.exe"

## WinGet-WrapperDetection.ps1
Matches locally installed version with newest available version using WinGet or specified exact version specified.  
Can be setup to accept newer installed version locally $AcceptNewerVersion

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/aac66723-24f9-4e7f-94ed-7a79ed49c623)

## WinGet-WrapperRequirements.ps1
Checks if application is detected locally. If not detected will not attempt update/install
To be used when only wanting to update if application is already installed. (Update Only)

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/1617c13e-11ef-4bfd-96c7-c4962706b0be)

Outputs either "Installed" or "Not Installed"

![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/b8cd24fd-da34-4e1c-aeb2-0627717e1244)

## Usage (InTune)
In InTune create an Windows app (Win32) and upload WinGet-Wrapper.InTuneWin as the package file.  
>**Install:** Powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -Context "System" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"  

>**Uninstall:** Powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -Context "System" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"

Change the $id variable to match the package id in the detection script and upload it  ($id = "VideoLAN.VLC")  
  *If specific version is required change the $TargetVersion (Ex. $TargetVersion = "1.0.0.0")*  
  *To define under which context to install set the $Context to either System or User*
  
![image](https://github.com/SorenLundt/WinGet-Wrapper/assets/127216441/6e29de05-effd-44e7-935b-1c3492d14af3)

Deploy the application and check log files in C:\ProgramData\WinGet-WrapperLogs
