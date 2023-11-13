# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Version History
# Version 1.1 - 27-02-2023 SorenLundt - Updated to retrieve current version from WinGet and dynamically check for latest version and detect based.
# Version 1.2 - 28-02-2023 SorenLundt - Added logging capability
# Version 1.3 - 01-03-2023 SorenLundt/ChatGPT - Changed expression to properly find the local installed version. Issue that winget list also shows available version from winget
# Version 1.4 - 01-03-2023 SorenLundt - Added section to cleanup old log files. 60 days.
# Version 1.5 - 13-03-2023 SorenLundt - Overall added better error handling - try/catch - Write-Log
# Version 1.6 - 13-03-2023 SorenLundt - Added functionality to report installed if locally installed version is newer than Available WinGet or TargetVersion. $AcceptNewerVersion
# Version 1.7 - 14-03-2023 SorenLundt - Added functionality to auto update - fixes issue #1 on https://github.com/SorenLundt/WinGet-Wrapper.
# Version 1.8 - 26-05-2023 SorenLundt - Added support for context choice by adding $Context variable
# Version 1.9 - 30-05-2023 SorenLundt - Fixed issues with running in user context (Winget path issues) + Updated version regex to be more precise
# Version 2.0 - 21-08-2023 SorenLundt - Removing AutoUpdate completely. Feature does not support InTune with "Available" deployments, as detection rules only runs periodically for "Required" deployments.
# Version 2.1 - 22-08-2023 SorenLundt - Adding UpdateOnly. To be used when only wanting to update apps but not install if not detected.
# Version 2.2 - 24-08-2023 SorenLundt - Removing UpdateOnly as this would lead to false positive, indicating that a given app was installed when actually it is not.
# Version 2.3 - 24-08-2023 SorenLundt - Adding automatically detection if running in user or system context. Removing Context parameter
# Version 2.4 - 24-08-2023 SorenLundt - WindowStyle Hidden for winget process + Other small fixes..
# Version 2.5 - 24-08-2023 SorenLundt - Added --scope $Context to winget cmd to avoid detecting applications in wrong context
# Version 2.6 - 18-10-2023 SorenLundt - Added --accept-source-agreements when searching for latest winget package to avoid any prompts
# Version 2.7 - 20-10-2023 SorenLundt - Fixed issues where applications containing + would not be detected.. Regex issue
# Version 2.8 - 23-10-2023 SorenLundt - Convert version string to System.Version objects to properly compare Winget and Installed versions
# Version 2.9 - 23-10-2023 SorenLundt - Updated version check segment + optimized detection by checking local installed version first
# Version 3.0 - 24-10-2023 SorenLundt - Fixed issue where packages containing + would not be able to search on winget and small minor changes.
# Version 3.1 - 27-10-2023 SorenLundt - Fixed issues with certain packages missing revision in version number, causing version mismatch compare to fail (ex. installed: 4.0.10  - Winget: 4.0.10.0)
# Version 3.2 - 13-11-2023 SorenLundt - Fixed issues with packages with build number (GitHub issue #6) Added function to correct empty(-1) major, minor, build, revision.  Sets it from -1 to 0
# Version 3.3 - 13-11-2023 SorenLundt - Added proper logging function instead of using Start-Transscript (Github Issue #5)
# Version 3.4 - 13-11-2023 SorenLundt - Minor issue. Wrong log filename, contained "Wrapper" instead of "Detection". Also now removing old *.txt files from log directory
# Version 3.5 - 13-11-2023 SorenLundt - Improved log output and added $ScriptVersion variable

# Settings
$id = "Exact WinGet Package ID" # WinGet Package ID - ex. VideoLAN.VLC
$TargetVersion = ""  # Set if specific version is desired (Optional)
$AcceptNewerVersion = "$True"   # Allows locally installed versions to be newer than $TargetVersion or available WinGet package version
# EndSettings

#Define common variables
$ScriptVersion = "3.5"
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
        [string]$LogFile = "$logPath\$($id)_WinGet_Detection_$($TimeStamp).log"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] $Message"

    # Output to the console
    Write-Host $LogEntry

    # Append to the log file
    $LogEntry | Out-File -Append -FilePath $LogFile
}

#Write useful variables to log
Write-Log "              WinGet-WrapperDetection v.$ScriptVersion"
Write-Log "https://github.com/SorenLundt/WinGet-Wrapper"
Write-Log "**************************************************"
Write-Log "Host: $env:ComputerName"
Write-Log "PackageName: $id"
Write-Log "TargetVersion: $TargetVersion"
Write-Log "AcceptNewerVersion = $AcceptNewerVersion"
Write-Log "LogPath: $logPath"
Write-Log "**************************************************"


# Clean log and txt files older than X days
$daysToKeepLogs = 60
$filesToDelete = Get-ChildItem $logPath -Recurse -Include *.log, *.txt | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$daysToKeepLogs)
try {
    $count = $filesToDelete.Count
    $filesToDelete | Remove-Item -Force | Out-Null
    if ($count -gt 0)
    {
    Write-Log "Cleaned up a total of $count old logs older than $daysToKeepLogs days."
    }
}
catch {
    Write-Log "Failed to delete old log files: $($_.Exception.Message)"
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

# Get latest version from WinGet
if ($TargetVersion -eq $null -or $TargetVersion -eq '') {
try {
    Start-Process -FilePath $wingetPath -ArgumentList "show --id $id --exact --accept-source-agreements" -WindowStyle Hidden -Wait -RedirectStandardOutput $stdout
    $winGetOutput = Get-Content -Path $stdout
    Remove-Item -Path $stdout -Force

    $TargetVersion = $winGetOutput | Select-String -Pattern "version:" | ForEach-Object { $_.Line -replace '.*version:\s*(.*)', '$1' }
    Write-Log "WinGet version: $TargetVersion"
}
catch {
    Write-Log "Failed to get latest version from WinGet: $($_.Exception.Message)"
    exit 1
}
}elseif ($TargetVersion -ne $null -and $TargetVersion -ne '')
{ Write-Host "TargetVersion: $TargetVersion (Set specific)"}


# Get version installed locally on machine
$InstalledVersion = $null  # Clear Variable
try {
    Start-Process -FilePath $wingetPath -ArgumentList "list $id --exact --accept-source-agreements --scope $Context" -WindowStyle Hidden -Wait -RedirectStandardOutput $stdout
    $searchString = Get-Content -Path $stdout
    Remove-Item -Path $stdout -Force

    # Remove + character from $id
    $searchString = $searchString -replace '\+', ''
    # Remove + character from $id
    $InstalledID = $id -replace '\+', ''

$versions = [regex]::Matches($searchString, "(?m)^.*$InstalledID\s*(?:[<>]?[\s]*)([\d.]+).*?$").Groups[1].Value
    if ($versions) {
        $InstalledVersion = ($versions | sort {[version]$_} | select -Last 1)
        Write-Log "Installed version: $InstalledVersion"
    }
    else {
        Write-Log "Package not found - #exit 1"
        exit 1
    }
} catch {
        Write-Log "Failed to get installed version: $($_.Exception.Message)"
        Write-Log "exit 1 - Report Not Installed"
        # Exit 1 - Report Not Installed$inst
        exit 1
    }


# Convert version strings to System.Version objects
$TargetVersion = [System.Version]::new($TargetVersion)
$InstalledVersion = [System.Version]::new($InstalledVersion)

# Function to set all version components to 0 if they are -1
function SetVersion($version) {
    $major = if ($version.Major -eq -1) { 0 } else { $version.Major }
    $minor = if ($version.Minor -eq -1) { 0 } else { $version.Minor }
    $build = if ($version.Build -eq -1) { 0 } else { $version.Build }
    $revision = if ($version.Revision -eq -1) { 0 } else { $version.Revision }
    return [System.Version]::new($major, $minor, $build, $revision)
}

$InstalledVersion = SetVersion $InstalledVersion
$TargetVersion = SetVersion $TargetVersion

# Check versions
if ($AcceptNewerVersion -eq $false -and $InstalledVersion -eq $TargetVersion) {
    Write-Log "exit 0 - Report Installed"
    exit 0
    # exit 0 - report installed 
}
elseif ($AcceptNewerVersion -eq $true -and $InstalledVersion -ge $TargetVersion) {
    Write-Log "exit 0 - Report Installed"
    exit 0
    # exit 0 - report installed 
}
else {
    Write-Log "exit 1 - Report Not Installed"
    exit 1
    # exit 1 - report not installed
}
