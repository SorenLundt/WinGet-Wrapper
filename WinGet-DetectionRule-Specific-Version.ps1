# Soren Lundt - 22-02-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Version History
# Version 1.1 - 27-02-2023 SOLU - Updated to retrieve current version from WinGet and dynamically check for latest version and detect based.
# Version 1.2 - 28-02-2023 SOLU - Added logging capability
# Version 1.3 - 01-03-2023 SOLU/ChatGPT - Changed expression to properly find the local installed version. Issue that winget list also shows available version from winget
# Version 1.4 - 01-03-2023 SOLU - Added section to cleanup old log files. 60 days.

#Define Package ID
$id = "Exact Package ID"
$TargetVersion = "1.0.0.0"

#Create log folder
if (!(Test-Path -Path $env:ProgramData\WinGet-WrapperLogs)) {
    New-Item -Path $env:ProgramData\WinGet-WrapperLogs -Force -ItemType Directory
}

#TimeStamp
$TimeStamp = "{0:MM-dd-yy}_{0:HH-mm-ss}" -f (Get-Date)

#Start Logging
Start-Transcript -Path "$env:ProgramData\WinGet-WrapperLogs\$($ID)_WinGet_Detection_$($TimeStamp).log"

#Clean log files older than X days
$Path = "$env:ProgramData\WinGet-WrapperLogs"
$Days = 60
$CutoffDate = (Get-Date).AddDays(-$Days)
$FilesToDelete = Get-ChildItem $Path -Recurse -Include *.log | Where-Object LastWriteTime -lt $CutoffDate
$Count = $FilesToDelete.Count
$FilesToDelete | Remove-Item -Force
Write-Host "Cleaned up a total of $Count old logs older than $Days days."

#Find WinGet.exe Location
$ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
if ($ResolveWingetPath){
    $WingetPath = $ResolveWingetPath[-1].Path
}
Set-Location $WingetPath

#Get version installed locally on machine
$searchstring = .\winget.exe list $id --exact --accept-source-agreements
$versions = [regex]::Matches($searchstring, "$id\s+([\d\.]+)").Groups[1].Value
if ($versions) {
    $InstalledVersion = ($versions | sort {[version]$_} | select -Last 1)
    Write-Host "Target Version:    $TargetVersion"
    Write-Host "Installed version: $InstalledVersion"
} else {
    Write-Host "Package not found"
}

# Check if version on WinGet and local machine matches
if ($TargetVersion -eq $InstalledVersion)
{
    Write-Output "Exit 0 - Report Installed"
    exit 0 # exit 0 - report installed
    
}
else
{
    Write-Output "Exit 1 - Report Not Installed"
    exit 1 # exit 1 - report not installed
}
