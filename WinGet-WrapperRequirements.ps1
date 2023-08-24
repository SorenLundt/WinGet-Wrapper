# Soren Lundt - 22-08-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Requirements script to check if desired application is installed. To be used when only wanting to update the application if already installed.  (UpdateOnly)
#
# Version History:
# Version 1.0 - 22-08-2023 SorenLundt - Initial version.
# Version 1.1 - 24-08-2023 SorenLundt - Adding automatically detection if running in user or system context. Removing Context parameter
# Version 1.2 - 24-08-2023 SorenLundt - WindowStyle Hidden for winget process + Other small fixes..

# Settings
$id = "VideoLAN.VLC" # WinGet Package ID - ex. VideoLAN.VLC

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
$logFile = "$logPath\$($id)_WinGet_Requirements_$($timestamp).log"
try {
    Start-Transcript -Path $logFile -Append
}
catch {
    Write-Output "Failed to start transcript: $($_.Exception.Message)"
    exit 1
}

#Write useful variables to log
Write-OutPut "ID = $id"

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

# Get version installed locally on machine
$InstalledVersion = $null  # Clear Variable
try {
    Start-Process -FilePath $wingetPath -ArgumentList "list $id --exact --accept-source-agreements" -WindowStyle Hidden -Wait -RedirectStandardOutput $stdout
    $searchString = Get-Content -Path $stdout
    Remove-Item -Path $stdout -Force
$versions = [regex]::Matches($searchString, "(?m)^.*$id\s*(?:[<>]?[\s]*)([\d.]+).*?$").Groups[1].Value

   
    if ($versions) {
        $InstalledVersion = ($versions | sort {[version]$_} | select -Last 1)
        Write-Output "Installed"
        exit 0
    }
    else {
        Write-Output "Not Installed"
        exit 0
    }
} catch {
        Write-Output "Not Installed"
        exit 0
    }