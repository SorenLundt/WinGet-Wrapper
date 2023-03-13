# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Version History
# Version 1.1 - 27-02-2023 SOLU - Updated to retrieve current version from WinGet and dynamically check for latest version and detect based.
# Version 1.2 - 28-02-2023 SOLU - Added logging capability
# Version 1.3 - 01-03-2023 SOLU/ChatGPT - Changed expression to properly find the local installed version. Issue that winget list also shows available version from winget
# Version 1.4 - 01-03-2023 SOLU - Added section to cleanup old log files. 60 days.
# Version 1.5 - 13-03-2023 SOLU - Overall added better error handling - try/catch - Write-error

# Define Package ID
$id = "Exact WinGet package ID"
$TargetVersion = "Version number if desired"  # Only set this if specific version is desired.

# Create log folder
$logPath = "$env:ProgramData\WinGet-WrapperLogs"
if (!(Test-Path -Path $logPath)) {
    try {
        New-Item -Path $logPath -Force -ItemType Directory | Out-Null
    }
    catch {
        Write-Error "Failed to create log directory: $($_.Exception.Message)"
        exit 1
    }
}

# Set up log file
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logFile = "$logPath\$($id)_WinGet_Detection_$($timestamp).log"
try {
    Start-Transcript -Path $logFile -Append
    Write-Output "Starting script"
}
catch {
    Write-Error "Failed to start transcript: $($_.Exception.Message)"
    exit 1
}

# Clean log files older than X days
$daysToKeep = 60
$filesToDelete = Get-ChildItem $logPath -Recurse -Include *.log | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$daysToKeep)
try {
    $count = $filesToDelete.Count
    $filesToDelete | Remove-Item -Force | Out-Null
    Write-Output "Cleaned up a total of $count old logs older than $daysToKeep days."
}
catch {
    Write-Error "Failed to delete old log files: $($_.Exception.Message)"
}

# Find WinGet.exe Location
$wingetPath = ""
try {
    $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
    if ($resolveWingetPath) {
        $wingetPath = $resolveWingetPath[-1].Path
        Set-Location $wingetPath
        Write-Output "WinGet path: $wingetPath"
    }
    else {
        Write-Error "Failed to find WinGet path"
        exit 1
    }
}
catch {
    Write-Error "Failed to find WinGet path: $($_.Exception.Message)"
    exit 1
}


# Get latest version from WinGet if dynamic is chosen
if ($TargetVersion -eq $null -and $TargetVersion -eq '') {
$TargetVersion = ""
try {
    $winGetOutput = .\winget.exe show --id "$id" --exact
    $TargetVersion = $winGetOutput | Select-String -Pattern "version:" | ForEach-Object { $_.Line -replace '.*version:\s*(.*)', '$1' }
    Write-Output "WinGet version: $TargetVersion"
}
catch {
    Write-Error "Failed to get latest version from WinGet: $($_.Exception.Message)"
    exit 1
}
}elseif ($TargetVersion -ne $null -and $TargetVersion -ne '')
{ Write-Host "TargetVersion: $TargetVersion (Set specific)"}

# Get version installed locally on machine
$InstalledVersion = ""
try {
    $searchString = .\winget.exe list "$id" --exact --accept-source-agreements
    $versions = [regex]::Matches($searchString, "$id\s+([\d\.]+)").Groups[1].Value
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
    Write-Error "Failed to get installed version: $($_.Exception.Message)"
    exit 1
}

# Check if version on WinGet and local machine matches
if ($InstalledVersion -eq $TargetVersion)
{
    Write-Output "Exit 0 - Report Installed"
    exit 0 # exit 0 - report installed
    
}
else
{
    Write-Output "Exit 1 - Report Not Installed"
    exit 1 # exit 1 - report not installed
}
