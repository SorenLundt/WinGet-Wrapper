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
System Context is not possible just by using "winget" as the .exe location must be found and this location is not static.

## Requirements
Newer Windows OS build that includes the WinGet natively in the OS.
Windows 10 20H2 or newer should be enough
## Limitations
* Only designed for System Context use  
  * Support for user context could be added later  
* WinGet-Detection-Dynamic-Version.ps1 can only be used with "Required" deployments in InTune - Issue: https://github.com/SorenLundt/WinGet-Wrapper/issues/1
  * An "Available" deployment would result in Company Portal seeing the application as not installed and would never update unless end-user click "Install" again. 


## WinGet-Wrapper.ps1
![image](https://user-images.githubusercontent.com/127216441/224036611-7bb907f9-7f26-42a1-b4ad-f4e95a1c930e.png)
#### Usage:
>.\WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -StopProcess "kill process using Stop-Process (.exe for the most part is not needed)" -ArgumentList "Arguments Passed to WinGet.exe"

## WinGet-WrapperDetection.ps1
Matches locally installed version with newest available version using WinGet or specified exact version specified.  
![image](https://user-images.githubusercontent.com/127216441/224702120-32020cc8-4b71-4348-a12c-473d24ca8857.png)

## Usage (InTune)
In InTune create an Windows app (Win32) and upload WinGet-Wrapper.InTuneWin as the package file.  
>**Install:** .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"  


>**Uninstall:** .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"

Change the $id variable to match the package id in the detection script and upload it  ($id = "VideoLAN.VLC")  
  *If specific version is required change the $TargetVersion (Ex. $TargetVersion = "1.0.0.0")*  
  
![image](https://user-images.githubusercontent.com/127216441/224702305-dca43b64-0d19-455f-bbcc-3a70e98933af.png)

Deploy the application and check log files in C:\ProgramData\WinGetLogs
