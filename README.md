# WinGet-Wrapper
A PowerShell to install/manage applications using WinGet - Common use case include InTune and ConfigMgr(SCCM)

# Background
WinGet have a few limitations in terms of automation and cannot easily be used with InTune and other endpoint management tools.

# Usage
Usage: .\WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -StopProcess "kill process using Stop-Process (.exe for the most part is not needed)" -ArgumentList "Arguments Passed to WinGet.exe"

#### Install example application using WinGet-Wrapper.ps1
>.\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"

#### Uninstall example application using WinGet-Wrapper.ps1
>.\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"


# Variables:
#### -PackageName 
Package name mainly used for naming the log file.

#### -StopProcess
Kill a specific process (Stop-process) before executing WinGet command 
(.exe should not be defined - uses Stop-Process) Skips any error automatically.

#### -ArgumentList
Arguments passed directly to WinGet.exe
