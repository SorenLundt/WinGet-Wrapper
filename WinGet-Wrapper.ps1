# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Wrapper for running winget in system context. Finds the WinGet install path and runs desired command
#
# Usage: .\WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -Context "User or System" -StopProcess "kill process using Stop-Process (.exe for the most part is not needed)" -ArgumentList "Arguments Passed to WinGet.exe"
# INSTALL   .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"
# UNINSTALL .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"
#
# Variables:
# PackageName = Package name mainly used for naming the log file.
# StopProcess = Kill a specific process (Stop-process) before executing WinGet command (.exe should not be defined) Skips any error automatically.
# ArgumentList = Arguments passed directly to WinGet
#
# Version History
# Version 1.0 - 22-02-2023 SorenLundt - Initial Version
# Version 1.1 - 28-02-2023 SorenLundt - Added logging functionality and changed to use Invoke-Expression to get data written to prompt and log file
# Version 1.2 - 01-03-2023 SorenLundt - Added logging of winget output + added PackageName parameter used in logging functionality
# Version 1.3 - 02-03-2023 SorenLundt - Added stop process to kill any process before upgrading. No user warning is shown etc.
# Version 1.4 - 26-05-2023 SorenLundt - Added support for context choice by adding $Context variable
# Version 1.5 - 30-05-2023 SorenLundt - Fixed issues with running in user context (Winget path issues)
# Version 1.6 - 24-08-2023 SorenLundt - Added automatically detection if running in user or system context. Decide context using deployment tools (InTune). Removing Context parameter
# Version 1.7 - 24-08-2023 SorenLundt - Added support for running pre/post script before winget action + WindowStyle Hidden for winget process + Other small fixes..

Param (
    # PackageName = Package name mainly used for naming the log file.
    [Parameter(Mandatory=$true)]
    [string]$PackageName,

    # StopProcess = Kill a specific process (Stop-process) before executing WinGet command (.exe should not be defined) Skips any error automatically.
    [Parameter()]
    [string]$StopProcess,

    # PreScript = Run a script before installation
    [Parameter()]
    [string]$PreScript,

    # PostScript = Run a script after installation
    [Parameter()]
    [string]$PostScript,
    
    # ArgumentList = Arguments passed directly to WinGet
    [Parameter(Mandatory=$true)]
    [string]$ArgumentList
)

#Define common variables
$logPath = "$env:ProgramData\WinGet-WrapperLogs"
$stdout = "$logPath\StdOut-$timestamp.txt"
$errout = "$logPath\ErrOut-$timestamp.txt"

# Create log folder
if (!(Test-Path -Path $logPath)) {
    try {
        New-Item -Path $logPath -Force -ItemType Directory | Out-Null
    }
    catch {
        Write-Output "Failed to create log directory: $($_.Exception.Message)"
        exit 1
    }
}

# Set up log file
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logFile = "$logPath\$($PackageName)_WinGet_Wrapper_$($TimeStamp).log"
try {
    Start-Transcript -Path $logFile -Append
}
catch {
    Write-Output "Failed to start transcript: $($_.Exception.Message)"
    exit 1
}

#Write useful variables to log
Write-Output "PackageName: $PackageName"
Write-Output "StopProcess: $StopProcess"
Write-Output "PreScript: $PreScript"
Write-Output "PostScript: $PostScript"
Write-Output "ArgumentList: $ArgumentList"

#PreScript
if (![string]::IsNullOrEmpty($PreScript)) {
    Write-OutPut "Running PreScript $PreScript"
    if (Test-Path $PreScript -PathType Leaf) {
        try {
            Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $PreScript
        } catch {
            Write-Output "An error occurred while running the PreScript: $_"
            Write-Output "Exit 1"
            exit 1
        }
    } else {
        Write-Output "PreScript not found: $PreScript"
        Write-Output "Exit 1"
        exit 1
    }
}

#Stop process
if (-not ($StopProcess -eq $null) -and $StopProcess -ne "") {
    Stop-Process -Name $StopProcess -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Write-Output "Stopped process: $StopProcess"
}

#Determine if running in system or user context
if ($env:USERNAME -like "*$env:COMPUTERNAME*") {
    Write-Output "Running in System Context"
    $Context = "System"
   }
   else {
    Write-Output "Running in User Context"
    $Context = "User"
   }

# Find WinGet.exe Location
if ($Context -contains "System"){
try {
    $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
    if ($resolveWingetPath) {
        $wingetPath = $resolveWingetPath[-1].Path
        Set-Location $wingetPath
        $wingetPath = $wingetPath + "\winget.exe"
        Write-Output "WinGet path: $wingetPath"
    }
    else {
        Write-Output "Failed to find WinGet path"
        exit 1
    }
}
catch {
    Write-Output "Failed to find WinGet path: $($_.Exception.Message)"
    exit 1
}
}
else{
    # Running in user context. Set WingetPath
$wingetPath = "winget.exe" 
}

Write-Output "Executing $wingetPath $ArgumentList"
Start-Process "$WingetPath" -ArgumentList "$ArgumentList" -WindowStyle Hidden -PassThru -Wait -RedirectStandardOutput "$stdout" -RedirectStandardError "$errout"

get-content "$stdout"
get-content "$errout"
Remove-item -Path "$stdout"
Remove-item -Path "$errout"

#PostScript
if (![string]::IsNullOrEmpty($PostScript)) {
    Write-OutPut "Running PostScript $PostScript"
    if (Test-Path $PostScript -PathType Leaf) {
        try {
            Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $PostScript
        } catch {
            Write-Output "An error occurred while running the PostScript: $_"
            Write-Output "Exit 1"
            exit 1
        }
    } else {
        Write-Output "PostScript not found: $PostScript"
        Write-Output "Exit 1"
        exit 1
    }
}

write-Output "Script Finished"