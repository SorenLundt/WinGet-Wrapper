# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Wrapper for running winget in system context. Finds the WinGet install path and runs desired command
#
# Usage: .\WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -StopProcess "kill process using Stop-Process (.exe for the most part is not needed)" -ArgumentList "Arguments Passed to WinGet.exe"
# INSTALL   .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"
# UNINSTALL .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"
#
# Variables:
# PackageName = Package name mainly used for naming the log file.
# StopProcess = Kill a specific process (Stop-process) before executing WinGet command (.exe should not be defined) Skips any error automatically.
# ArgumentList = Arguments passed directly to WinGet
#
# Version History
# Version 1.0 - 22-02-2023 SOLU - Initial Version
# Version 1.1 - 28-02-2023 SOLU - Added logging functionality and changed to use Invoke-Expression to get data written to prompt and log file
# Version 1.2 - 01-03-2023 SOLU - Added logging of winget output + added PackageName parameter used in logging functionality
# Version 1.3 - 02-03-2023 SOLU - Added stop process to kill any process before upgrading. No user warning is shown etc.

Param (
    [Parameter(Mandatory=$true)]
    [string]$PackageName,

    [Parameter()]
    [string]$StopProcess,

    [Parameter(Mandatory=$true)]
    [string]$ArgumentList
)

  #Create log folder
if (!(Test-Path -Path $env:ProgramData\WinGet-WrapperLogs)) {
    New-Item -Path $env:ProgramData\WinGet-WrapperLogs -Force -ItemType Directory
}

#TimeStamp
$TimeStamp = "{0:MM-dd-yy}_{0:HH-mm-ss}" -f (Get-Date)
Write-Host "PackageName: $PackageName"
Write-Host "ArgumentList: $Arguments"
Write-Host "Timestamp: $TimeStamp"

#Start Logging
Start-Transcript -Path "$env:ProgramData\WinGet-WrapperLogs\$($PackageName)_WinGet_Wrapper_$($TimeStamp).log"

#Stop process
if (-not ($StopProcess -eq $null) -and $StopProcess -ne "") {
    Stop-Process -Name $StopProcess -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Write-Host "Stopped process: $StopProcess"
}

$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
   if ($ResolveWingetPath){
           $WingetPath = $ResolveWingetPath[-1].Path
    }
Set-Location $wingetPath
$stdout = "$env:ProgramData\WinGet-WrapperLogs\StdOut-$timestamp.txt"
$errout = "$env:ProgramData\WinGet-WrapperLogs\ErrOut-$timestamp.txt"

Write-Host "Executing $wingetPath\Winget.exe $ArgumentList"
Start-Process "$WingetPath\winget.exe" -ArgumentList "$ArgumentList" -PassThru -Wait -RedirectStandardOutput "$stdout" -RedirectStandardError "$errout"

get-content "$stdout"
get-content "$errout"
Remove-item -Path "$stdout"
Remove-item -Path "$errout"
write-host "Script Finished"