# WinGet-Wrapper  
A PowerShell to install/manage applications using WinGet - Common use case include InTune, and other endpoint management products.´ 
Detection script available for fixed version and for dynamic version (automatically matches the installed version with version available with WinGet)

* Dynamically finds the WinGet directory to be used under System Context.   
* Kill selected process before WinGet command using -StopProcess.  
* Detection script that dynamically finds latest package available trough WinGet
* Logs to $env:ProgramData\WinGet-WrapperLogs (Usually C:\ProgramData\WinGet-WrapperLogs) 
* Detection script performs automatic cleanup of log files older than 60 days.

## Background  
WinGet have a few limitations in terms of automation and is not integrated with common endpoints management products.  
System Context is not possible just by using "winget" as the .exe location must be found and this location is not static due to versioning in the directory name.

## Requirements
Newer Windows OS build that includes the WinGet natively in the OS.
Windows 10 20H2 or newer should be enough
## Limitations
* Only designed for System Context use  
  * Support for user context could be added later  

## WinGet-Wrapper.ps1
![image](https://user-images.githubusercontent.com/127216441/224036611-7bb907f9-7f26-42a1-b4ad-f4e95a1c930e.png)
#### Usage:
>.\WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -StopProcess "kill process using Stop-Process (.exe for the most part is not needed)" -ArgumentList "Arguments Passed to WinGet.exe"

## WinGet-WrapperDetection.ps1
Matches locally installed version with newest available version using WinGet or specified exact version specified.  
![image](https://user-images.githubusercontent.com/127216441/225035005-5d2a7860-4178-43b6-855e-20db6b33a38f.png)

Example:
$Target Version =

## Usage (InTune)
In InTune create an Windows app (Win32) and upload WinGet-Wrapper.InTuneWin as the package file.  
>**Install:** .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"  

>**Uninstall:** .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"

Change the $id variable to match the package id in the detection script and upload it  ($id = "VideoLAN.VLC")  
  *If specific version is required change the $TargetVersion (Ex. $TargetVersion = "1.0.0.0")*  
  *If AutoUpdate is desired define the $AutoUpdate variable*  
  *If AutoUpdate is desired define the $AutoUpdateArgumentList variable*  
  *If StopProcess when using AutoUpdate is desired define the $AutoUpdateStopProcess variable"*  
  
![image](https://user-images.githubusercontent.com/127216441/225034075-44ae1fe6-25db-49fe-8655-7d94be8584c3.png)

Deploy the application and check log files in C:\ProgramData\WinGet-WrapperLogs
