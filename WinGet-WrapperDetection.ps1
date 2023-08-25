# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Version History
# Version 1.1 - 27-02-2023 SorenLundt - Updated to retrieve current version from WinGet and dynamically check for latest version and detect based.
# Version 1.2 - 28-02-2023 SorenLundt - Added logging capability
# Version 1.3 - 01-03-2023 SorenLundt/ChatGPT - Changed expression to properly find the local installed version. Issue that winget list also shows available version from winget
# Version 1.4 - 01-03-2023 SorenLundt - Added section to cleanup old log files. 60 days.
# Version 1.5 - 13-03-2023 SorenLundt - Overall added better error handling - try/catch - Write-Output
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

# Settings
$id = "Exact WinGet Package ID" # WinGet Package ID - ex. VideoLAN.VLC
$TargetVersion = ""  # Set if specific version is desired (Optional)
$AcceptNewerVersion = $True   # Allows locally installed versions to be newer than $TargetVersion or available WinGet package version

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
        Write-Output "Failed to create log directory: $($_.Exception.Message)"
        exit 1
    }
}

# Set up log file
$logFile = "$logPath\$($id)_WinGet_Detection_$($timestamp).log"
try {
    Start-Transcript -Path $logFile -Append
}
catch {
    Write-Output "Failed to start transcript: $($_.Exception.Message)"
    exit 1
}

#Write useful variables to log
Write-Output "ID: $id"
Write-Output "TargetVersion: $TargetVersion"
Write-OutPut "AcceptNewerVersion = $AcceptNewerVersion"

# Clean log files older than X days
$daysToKeepLogs = 60
$filesToDelete = Get-ChildItem $logPath -Recurse -Include *.log | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$daysToKeepLogs)
try {
    $count = $filesToDelete.Count
    $filesToDelete | Remove-Item -Force | Out-Null
    Write-Output "Cleaned up a total of $count old logs older than $daysToKeepLogs days."
}
catch {
    Write-Output "Failed to delete old log files: $($_.Exception.Message)"
}

#Determine if running in system or user context
if ($env:USERNAME -like "*$env:COMPUTERNAME*") {
    Write-Output "Running in System Context"
    $Context = "Machine"
   }
   else {
    Write-Output "Running in User Context"
    $Context = "User"
   }

# Find WinGet.exe Location
if ($Context -contains "Machine"){
    try {
        $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
        if ($resolveWingetPath) {
            $wingetPath = $resolveWingetPath[-1].Path
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

# Get latest version from WinGet
if ($TargetVersion -eq $null -or $TargetVersion -eq '') {
try {
    Start-Process -FilePath $wingetPath -ArgumentList "show --id $id --exact" -WindowStyle Hidden -Wait -RedirectStandardOutput $stdout
    $winGetOutput = Get-Content -Path $stdout
    Remove-Item -Path $stdout -Force

    $TargetVersion = $winGetOutput | Select-String -Pattern "version:" | ForEach-Object { $_.Line -replace '.*version:\s*(.*)', '$1' }
    Write-Output "WinGet version: $TargetVersion"
}
catch {
    Write-Output "Failed to get latest version from WinGet: $($_.Exception.Message)"
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
$versions = [regex]::Matches($searchString, "(?m)^.*$id\s*(?:[<>]?[\s]*)([\d.]+).*?$").Groups[1].Value

    if ($versions) {
        $InstalledVersion = ($versions | sort {[version]$_} | select -Last 1)
        Write-Output "Installed version: $InstalledVersion"
    }
    else {
        Write-Output "Package not found - #exit 1"
        exit 1
    }
} catch {
        Write-Output "Failed to get installed version: $($_.Exception.Message)"
        # Exit 1 - Report Not Installed
        exit 1
    }

if ($InstalledVersion -ge $TargetVersion -or ($AcceptNewerVersion -and $InstalledVersion -gt $TargetVersion)){
    Write-Output "exit 0 - Report Installed"
    exit 0 #exit 0 - report installed 
}
else {
    Write-Output "exit 1 - Report Not Installed"
    exit 1 #exit 1 - report not installed
}