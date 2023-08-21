# Soren Lundt - 21-08-2023 - https://github.com/SorenLundt/WinGet-Wrapper
#
# Purpose:
# Use as detection script for required deployments to update applications if installed. 
# If application is not installed initially, it will not install it.
#
# Version History:
# Version 1.0 - 21-08-2023 SOLU - Inital version "Winget-WrapperUpdateDetection" 
#


# Settings
$id = "VideoLAN.VLC" # WinGet Package ID - ex. VideoLAN.VLC
$TargetVersion = ""  # Set if specific version is desired (Optional)
$AcceptNewerVersion = $True   # Allows locally installed versions to be newer than $TargetVersion or available WinGet package version
$Context = "System" # Set to either System or User

# Create log folder
$logPath = "$env:ProgramData\WinGet-WrapperLogs"
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
$logFile = "$logPath\$($id)_WinGet_Detection_$($timestamp).log"
try {
    Start-Transcript -Path $logFile -Append
}
catch {
    Write-Output "Failed to start transcript: $($_.Exception.Message)"
    exit 1
}

#Write useful variable settings to log
Write-OutPut "ID = $id"
Write-OutPut "TargetVersion = $TargetVersion  (Blank = Latest Version)"
Write-OutPut "AcceptNewerVersion = $AcceptNewerVersion"
Write-OutPut "Context = $Context"



# Clean log files older than X days
$daysToKeep = 60
$filesToDelete = Get-ChildItem $logPath -Recurse -Include *.log | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$daysToKeep)
try {
    $count = $filesToDelete.Count
    $filesToDelete | Remove-Item -Force | Out-Null
    Write-Output "Cleaned up $count old logs older than $daysToKeep days."
}
catch {
    Write-Output "Failed to delete old log files: $($_.Exception.Message)"
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
Write-Output "Running in User Context"
$wingetPath = "winget.exe" 
}

# Get latest version from WinGet if dynamic is chosen
if ($TargetVersion -eq $null -or $TargetVersion -eq '') {
try {
    if ($Context -contains "System"){
    $winGetOutput = .\winget.exe show --id "$id" --exact
    }
    else {
    $winGetOutput = winget.exe show --id "$id" --exact
    }
    $TargetVersion = $winGetOutput | Select-String -Pattern "version:" | ForEach-Object { $_.Line -replace '.*version:\s*(.*)', '$1' }
    Write-Output "WinGet version: $TargetVersion"
}
catch {
    Write-Output "Failed to get latest version from WinGet: $($_.Exception.Message)"
    exit 1
}
}elseif ($TargetVersion -ne $null -and $TargetVersion -ne '')
{ Write-Host "TargetVersion: $TargetVersion (Set specific)"}

# Write out Application Name

# Get version installed locally on machine
$InstalledVersion = $null  # Clear Variable
try {
    if ($Context -contains "System"){
    $searchString = .\winget.exe list "$id" --exact --accept-source-agreements
    }
    else {
    $searchString = winget.exe list "$id" --exact --accept-source-agreements
    }
$versions = [regex]::Matches($searchString, "(?m)^.*$id\s*(?:[>]?[\s]*)([\d.]+).*?$").Groups[1].Value

   
    if ($versions) {
        $InstalledVersion = ($versions | sort {[version]$_} | select -Last 1)
        Write-Output "Installed version: $InstalledVersion"
    }
    else {
        Write-Output "Package not found - exit 1"
        exit 1
    }
}
catch {
    Write-Output "Application '$id' not installed. Nothing to update"
    exit 0
}


if ($InstalledVersion -ge $TargetVersion -or ($AcceptNewerVersion -and $InstalledVersion -gt $TargetVersion)){
    Write-Output "Exit 0 - Report Installed"
    exit 0 # exit 0 - report installed 
}
else {
    Write-Output "Exit 1 - Report Not Installed"
    exit 1 # exit 1 - report not installed
}