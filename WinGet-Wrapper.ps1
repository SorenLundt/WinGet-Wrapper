# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Wrapper for running winget in system context. Finds the WinGet install path and runs desired command
#
# Usage: .\WinGet-Wrapper.ps1 -PackageName "Packagename used in log entry" -Context "User or System" -StopProcess "kill process using Stop-Process (.exe for the most part is not needed)" -ArgumentList "Arguments Passed to WinGet.exe"
# INSTALL   .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -PreScript "Script1.ps1" -PostScript "Script2.ps1" -ArgumentList "install --exact --id VideoLAN.VLC --silent --accept-package-agreements --accept-source-agreements --scope machine"
# UNINSTALL .\WinGet-Wrapper.ps1 -PackageName "VideoLAN.VLC" -StopProcess "VLC" -PreScript "Script1.ps1" -PostScript "Script2.ps1" -ArgumentList "Uninstall --exact --id VideoLAN.VLC --silent --accept-source-agreements --scope machine"
#
# Variables:
# PackageName = Package name mainly used for naming the log file.
# StopProcess = Kill a specific process (Stop-process) before executing WinGet command (.exe should not be defined) Skips any error automatically.
# PreScript = Run powershell script before installation
# PostScript = Run powershell script after installation
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
# Version 1.8 - 13-11-2023 SorenLundt - Added proper logging function instead of using Start-Transscript (Github Issue #5)
# Version 1.9 - 13-11-2023 SorenLundt - Improved log output and added $ScriptVersion variable
$ScriptVersion = "1.9"

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
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = "$env:ProgramData\WinGet-WrapperLogs"
$stdout = "$logPath\StdOut-$timestamp.txt"
$errout = "$logPath\ErrOut-$timestamp.txt"

# Create log folder
if (!(Test-Path -Path $logPath)) {
    try {
        New-Item -Path $logPath -Force -ItemType Directory | Out-Null
    }
    catch {
        Write-Log "Failed to create log directory: $($_.Exception.Message)"
        exit 1
    }
}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "$logPath\$($PackageName)_WinGet_Wrapper_$($TimeStamp).log"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] $Message"

    # Output to the console
    Write-Host $LogEntry

    # Append to the log file
    $LogEntry | Out-File -Append -FilePath $LogFile
}


#Write useful variables to log
Write-Log "                    WinGet-Wrapper v.$ScriptVersion"
Write-Log "https://github.com/SorenLundt/WinGet-Wrapper"
Write-Log "**************************************************"
Write-Log "Host: $env:ComputerName"
Write-Log "PackageName: $PackageName"
Write-Log "StopProcess: $StopProcess"
Write-Log "PreScript: $PreScript"
Write-Log "PostScript: $PostScript"
Write-Log "ArgumentList: $ArgumentList"
Write-Log "LogPath: $logPath"
Write-Log "**************************************************"

#PreScript
if (![string]::IsNullOrEmpty($PreScript)) {
    Write-Log "Running PreScript $PreScript"
    if (Test-Path $PreScript -PathType Leaf) {
        try {
            Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $PreScript
        } catch {
            Write-Log "An error occurred while running the PreScript: $_"
            Write-Log "Exit 1"
            exit 1
        }
    } else {
        Write-Log "PreScript not found: $PreScript"
        Write-Log "Exit 1"
        exit 1
    }
}

#Stop process
if (-not ($StopProcess -eq $null) -and $StopProcess -ne "") {
    Stop-Process -Name $StopProcess -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Write-Log "Stopped process: $StopProcess"
}

#Determine if running in system or user context
if ($env:USERNAME -like "*$env:COMPUTERNAME*") {
    Write-Log "Running in System Context"
    $Context = "Machine"
   }
   else {
    Write-Log "Running in User Context"
    $Context = "User"
   }

# Find WinGet.exe Location
if ($Context -contains "Machine"){
try {
    $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
    if ($resolveWingetPath) {
        $wingetPath = $resolveWingetPath[-1].Path
        $wingetPath = $wingetPath + "\winget.exe"
        Write-Log "WinGet path: $wingetPath"
    }
    else {
        Write-Log "Failed to find WinGet path"
        exit 1
    }
}
catch {
    Write-Log "Failed to find WinGet path: $($_.Exception.Message)"
    exit 1
}
}
else{
    # Running in user context. Set WingetPath
$wingetPath = "winget.exe" 
}

Write-Log "Executing $wingetPath $ArgumentList"
Start-Process "$WingetPath" -ArgumentList "$ArgumentList" -WindowStyle Hidden -PassThru -Wait -RedirectStandardOutput "$stdout" -RedirectStandardError "$errout"

# Log the contents of the stdout and errout
$stdoutContent = Get-Content "$stdout"
$erroutContent = Get-Content "$errout"

# Log the contents of the stdout and errout  (single lines)
$stdoutContent -split "`r`n" | ForEach-Object {
    Write-Log "  $_"
}
$erroutContent -split "`r`n" | ForEach-Object {
    Write-Log "  $_"
}


Remove-item -Path "$stdout"
Remove-item -Path "$errout"

#PostScript
if (![string]::IsNullOrEmpty($PostScript)) {
    Write-Log "Running PostScript $PostScript"
    if (Test-Path $PostScript -PathType Leaf) {
        try {
            Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $PostScript
        } catch {
            Write-Log "An error occurred while running the PostScript: $_"
            Write-Log "Exit 1"
            exit 1
        }
    } else {
        Write-Log "PostScript not found: $PostScript"
        Write-Log "Exit 1"
        exit 1
    }
}

Write-Log "Script Finished"