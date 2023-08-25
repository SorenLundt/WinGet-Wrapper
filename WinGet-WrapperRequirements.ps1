# Soren Lundt - 22-08-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Requirements script to check if desired application is installed. To be used when only wanting to update the application if already installed.  (UpdateOnly)
#
# Version History:
# Version 1.0 - 22-08-2023 SorenLundt - Initial version.
# Version 1.1 - 24-08-2023 SorenLundt - Adding automatically detection if running in user or system context. Removing Context parameter
# Version 1.2 - 24-08-2023 SorenLundt - WindowStyle Hidden for winget process + Other small fixes..
# Version 1.3 - 25-08-2023 SorenLundt - Removing logging part. The script must only output "Installed" or "Not Installed"

# Settings
$id = "Exact WinGet Package ID" # WinGet Package ID - ex. VideoLAN.VLC

#Define common variables
$logPath = "$env:ProgramData\WinGet-WrapperLogs"
$stdout = "$logPath\StdOut-$timestamp.txt"
$errout = "$logPath\ErrOut-$timestamp.txt"

#Determine if running in system or user context
if ($env:USERNAME -like "*$env:COMPUTERNAME*") {
    $Context = "System"
   }
   else {
    $Context = "User"
   }

# Find WinGet.exe Location
if ($Context -contains "System"){
    try {
        $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
        if ($resolveWingetPath) {
            $wingetPath = $resolveWingetPath[-1].Path
            $wingetPath = $wingetPath + "\winget.exe"
            #Write-Output "WinGet path: $wingetPath"
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