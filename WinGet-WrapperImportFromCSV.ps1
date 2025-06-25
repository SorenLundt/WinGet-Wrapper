# Soren Lundt - 12-09-2023
# URL: https://github.com/SorenLundt/WinGet-Wrapper
# License: https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/LICENSE.txt
#
# Imports packages from WinGet to InTune (incuding available WinGet package metadata)
# Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm
# Logs file under Logs\WinGet_WrapperImportFromCSV_yyyy-MM-dd_HH-mm-ss.log
# 
# Usage Example: .\WinGet-WrapperImportFromCSV.ps1 -TenantID company.onmicrosoft.com -csvFile WinGet-WrapperImportFromCSV.csv -SkipConfirmation
#
# Parameters:
# csvFile = csvFile to import from (default: WinGet-WrapperImportFromCSV.csv)
# TenantID = TenantID to connect to MSGraph/InTune
# LogFile = Manually define logfile path. If not defined, default will be used
# ScriptRoot = Do not set. Define script root path. Only used when running import job from WinGet-WrapperImportGUI
# SkipConfirmation = Skips confirmation for each package
# SkipModuleCheck = Do not install or update required modules.
# 
#
# csvFile columns:
# PackageID,Context,AcceptNewerVersion,UpdateOnly,TargetVersion,StopProcessInstall,StopProcessUninstall,PreScriptInstall,PostScriptInstall,PreScriptUninstall,PostScriptUninstall,CustomArgumentListInstall,CustomArgumentListUninstall,InstallIntent,Notification,GroupID
#
# Requirements:
# Requires Script files and IntuneWinAppUtil.exe to be present in script directory
#
# Version History
# Version 1.0 - 12-09-2023 SorenLundt - Initial Version
# Version 1.1 - 14-09-2023 SorenLundt - Replaced ' with [char]34 (Quotation mark)  InTune does not handle ' well
# Version 1.2 - 20-09-2023 SorenLundt - Added possiblity to deploy application once imported. Set via CSV file (InstallIntent, Notification, GroupID)
# Version 1.3 - 20-09-2023 SorenLundt - If -SkipConfirmation set will skip the entire WinGet package output section
# Version 1.4 - 27-09-2023 SorenLundt - Added -OverWrite which will delete apps with the same display name
# Version 1.5 - 27-09-2023 SorenLundt - Remove -OverWrite feature - Requires more work.. Some bugs
# Version 1.6 - 24-10-2023 SorenLundt - When importing packages with UpdateOnly=1 will now use "winget update" instead of "winget install"
# Version 1.7 - 13-11-2023 SorenLundt - Minor script comment and code mismatch for replacing AcceptNeverVersion with $False (was $True)  Github issue #7
# Version 1.8 - 13-11-2023 SorenLundt - Removed check if package already exists in InTune. Not reliable. Needs work. Improved required PS modules check/installation
# Version 1.9 - 14-11-2023 SorenLundt - Fixed issue where $True was not being replace by $False for $Row.AcceptNewerVersion value - Github issue #7
# Version 2.0 - 28-11-2023 SorenLundt - Changed various file paths to be able to use script from GUI + added logging
# Version 2.1 - 12-02-2024 SorenLundt - Various improvements/changes to support usage from WinGet-WrapperImportGUI (-SkipModuleCheck -Scriptroot)
# Version 2.2 - 19-02-2023 SorenLundt - Added --accept-source-agreements to winget show commands, to avoid prompt - Github issue #12

#Parameters
Param (
    #CSV File to import from (default: WinGet-WrapperImportFromCSV.csv)
    [Parameter()]
    [string]$csvFile = "$PSScriptRoot\WinGet-WrapperImportFromCSV.csv",
    
    #ClientID to connect to MSGraph/InTune with Connect-MSIntuneGraph
    [Parameter(Mandatory = $False)]
    [string]$ClientID = "14d82eec-204b-4c2f-b7e8-296a70dab67e",

    #Redirect URI of Enterprise application
    [Parameter(Mandatory = $False)]
    [string]$RedirectURL = "https://login.microsoftonline.com/common/oauth2/nativeclient",
    
    #TenantID to connect to MSGraph/InTune
    [Parameter(Mandatory = $True)]
    [string]$TenantID,

    #LogFile (Manually define logfile path. If not defined, default will be used)
    [string]$LogFile,

    #ScriptRoot (Do not set. Define script root path. Only used when running import job from WinGet-WrapperImportGUI)
    [string]$ScriptRoot,

    #Skips confirmation for each package before import
    [Switch]$SkipConfirmation = $false,
    
    #Skips Module check
    [Switch]$SkipModuleCheck = $false
)

#Timestamp
$TimeStamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'

# Default Log File
if (-not $LogFile) {
    $LogFolder = Join-Path -Path $PSScriptRoot -ChildPath 'Logs'

    # Create logs folder if it doesn't exist
    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory | Out-Null
    }

    $LogFile = Join-Path -Path $LogFolder -ChildPath "WinGet_WrapperImportFromCSV_$TimeStamp.log"
}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "$LogFile"
    )

    $LogMsgTimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogEntry = "[$LogMsgTimeStamp] $Message"

    # Output to the console
    Write-Host $LogEntry

    # Append to the log file
    $LogEntry | Out-File -Append -FilePath $LogFile
}
Write-Log "Log File: $LogFile"

Write-log "Parameters: csvFile: $csvFile, TenantID: $TenantID, SkipConfirmation: $SkipConfirmation, LogFile: $LogFile, SkipModuleCheck: $SkipModuleCheck, scriptRoot: $scriptRoot"

# Install and load required modules
# https://github.com/MSEndpointMgr/IntuneWin32App
if (-not $SkipModuleCheck) {
    $intuneWin32AppModule = 'IntuneWin32App'
    $microsoftGraphIntuneModule = 'Microsoft.Graph.Intune'

    # Check if update is required for IntuneWin32App module
    $moduleInstalled = Get-InstalledModule -Name $intuneWin32AppModule -ErrorAction SilentlyContinue

    if (-not $moduleInstalled) {
        Install-Module -Name $intuneWin32AppModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $intuneWin32AppModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $intuneWin32AppModule -Force
        }
        else {
            Write-Host "Module $intuneWin32AppModule is already up-to-date."
        }
    }

    # Check if update is required for Microsoft.Graph.Intune module
    $moduleInstalled = Get-InstalledModule -Name $microsoftGraphIntuneModule -ErrorAction SilentlyContinue

    if (-not $moduleInstalled) {
        Install-Module -Name $microsoftGraphIntuneModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $microsoftGraphIntuneModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $microsoftGraphIntuneModule -Force
        }
        else {
            Write-Host "Module $microsoftGraphIntuneModule is already up-to-date."
        }
    }
}
#Import modules
Import-Module -Name 'IntuneWin32App'
Import-Module -Name 'Microsoft.Graph.Intune'

# Welcome greeting
Write-Log ' '
Write-Log ' '
Write-Log '-----------------------------'
Write-Log '---- WinGet-WrapperCreate----'
Write-Log '-----------------------------'
Write-Log ' '
Write-Log ' '
Write-Log 'https://github.com/SorenLundt/WinGet-Wrapper'
Write-Log '       GNU General Public License v3'

# Test CSV path
if (Test-Path -Path "$csvFile" -PathType Leaf) {
    Write-Log "File: $csvFile"
}
else {
    Write-Log "File not found: $csvFile" -ForegroundColor 'Red'
    return
}

# Import the CSV file with custom headers
$data = Import-Csv -Path "$csvFile" -Header 'PackageID', 'Context', 'AcceptNewerVersion', 'UpdateOnly', 'TargetVersion', 'StopProcessInstall', 'StopProcessUninstall', 'PreScriptInstall', 'PostScriptInstall', 'PreScriptUninstall', 'PostScriptUninstall', 'CustomArgumentListInstall', 'CustomArgumentListUninstall', 'InstallIntent', 'Notification', 'GroupID' | Select-Object -Skip 1

# Convert "AcceptNewerVersion" and "UpdateOnly" columns to Boolean values
$data = $data | ForEach-Object {
    $_.AcceptNewerVersion = [bool]($_.AcceptNewerVersion -as [int])
    $_.UpdateOnly = [bool]($_.UpdateOnly -as [int])
    $_
}
Write-Log '-- IMPORT LIST --'
foreach ($row in $data) {
    Write-Log "IMPORT PackageID:$($row.PackageID) - Context:$($row.Context) - UpdateOnly:$($row.UpdateOnly) - TargetVersion:$($row.TargetVersion)" -ForegroundColor Gray
}
Write-Log ''

Write-Log '-- DEPLOY LIST --'
foreach ($row in $data) {
    if ($null = $row.GroupID -or $row.GroupID -ne '') {
        Write-Log "DEPLOY PackageID:$($row.PackageID) GroupID:$($row.GroupID) InstallIntent:$($row.InstallIntent) Notification:$($row.Notification)" -ForegroundColor Gray
    }
}

#Connect to Intune
#if (-not $SkipInTuneConnection){
try {  
    Connect-MSIntuneGraph -TenantID "$TenantID" -ClientID $ClientID -RedirectUri $RedirectURL -Interactive
}
catch {
    Write-Log 'ERROR: Connect-MSIntuneGraph Failed. Exiting' -ForegroundColor 'Red'
    break
}
#}


#Import each application to InTune
foreach ($row in $data) {
    & {
        try {
            #Write-Log "--- Package Details ---"
            #Write-Log "PackageID: $($row.PackageID)"
            #Write-Log "Context: $($row.Context)"
            #Write-Log "AcceptNewerVersion: $($row.AcceptNewerVersion)"
            #Write-Log "UpdateOnly: $($row.UpdateOnly)"
            #Write-Log "TargetVersion: $($row.TargetVersion)"
            #Write-Log "StopProcessInstall: $($row.StopProcessInstall)"
            #Write-Log "StopProcessUninstall: $($row.StopProcessUninstall)"
            #Write-Log "PreScriptInstall: $($row.PreScriptInstall)"
            #Write-Log "PostScriptInstall: $($row.PostScriptInstall)"
            #Write-Log "PreScriptUninstall: $($row.PreScriptUninstall)"
            #Write-Log "PostScriptUninstall: $($row.PostScriptUninstall)"
            #Write-Log "CustomArgumentListInstall: $($row.CustomArgumentListInstall)"
            #Write-Log "CustomArgumentListInstall: $($row.CustomArgumentListUninstall)"

            #TimeStamp
            $Timestamp = (Get-Date).ToString('yyyyMMddHHmmss')

            #Get User
            $currentUser = $env:USERNAME

            Write-Log '--- Validation Start ---'
            Write-Log "Validating Package $($row.PackageID)"
            #Check context is valid
            if ($row.Context -notin @('Machine', 'machine', 'User', 'user')) {
                Write-Log "Invalid context setting $($row.Context) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor 'Red'
                break
            }

            #Check AcceptNewerVersion is true or false
            if ($row.AcceptNewerVersion -ne $True -and $row.AcceptNewerVersion -ne $False) {
                Write-Log "Invalid AcceptNewerVersion setting $($row.AcceptNewerVersion) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor 'Red'
                break
            }

            #Check InstallIntent is valid
            if ($null -ne $row.InstallIntent -and $row.InstallIntent -ne '') {
                if ($row.InstallIntent -notcontains 'Required' -and $row.InstallIntent -notcontains 'required' -and $row.InstallIntent -notcontains 'Available' -and $row.InstallIntent -notcontains 'available') {
                    Write-Log "Invalid InstallIntent setting $($row.InstallIntent) for package $($row.PackageID) found in CSV. Please review CSV (Use Available or Required) Exiting.." -ForegroundColor 'Red'
                    break
                }
            }

            #Check Notification is valid
            if ($null -ne $row.Notification -and $row.Notification -ne '') {
                $validNotificationValues = 'showAll', 'showReboot', 'hideAll'
                if ($validNotificationValues -notcontains $row.Notification.ToLower()) {
                    Write-Log "Invalid Notification setting $($row.Notification) for package $($row.PackageID) found in CSV. Please review CSV (Use showAll, showReboot, or hideAll) Exiting.." -ForegroundColor 'Red'
                    break
                }
            }

            #Check GroupID is set if InstallIntent is specified
            if ($null -ne $row.InstallIntent -and $row.InstallIntent -ne '') {
                if ($row.GroupID -eq $null -or $row.GroupID -eq '') {
                    Write-Log "Invalid GroupID setting $($row.GroupID) for package $($row.PackageID) found in CSV. Please review CSV. If InstallIntent is set, GroupID must be set too!"
                    break
                }
            }

            #Check UpdateOnly is true or false
            Write-Log "Checking UpdateOnly Value for $($row.PackageID)"
            if ($row.UpdateOnly -ne $True -and $row.UpdateOnly -ne $False) {
                Write-Log "Invalid UpdateOnly setting $($row.UpdateOnly) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor 'Red'
                break
            }

            #Check StopProcessInstall and StopProcessUninstall does not contain "exe"  (inform should not contain .exe)
            Write-Log "Checking StopProcessInstall Value for $($row.PackageID)"
            if ($row.StopProcessInstall -contains '.') {
                Write-Log "Invalid StopProcessInstall setting $($row.StopProcessInstall) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor 'Red'
                break
            }
            Write-Log "Checking StopProcessUninstall Value for $($row.PackageID)"
            if ($row.StopProcessUninstall -contains '.') {
                Write-Log "Invalid StopProcessUninstall setting $($row.StopProcessUninstall) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor 'Red'
                break
            }

            #Check PreScriptInstall, PostScriptInstall, PreScriptUninstall, PostScriptUninstall contains .ps1  (inform must be .ps1)
            <#
if (
    ($null -ne $row.PreScriptInstall -and -not [string]::IsNullOrEmpty($row.PreScriptInstall)) -or
    ($null -ne $row.PostScriptInstall -and -not [string]::IsNullOrEmpty($row.PostScriptInstall)) -or
    ($null -ne $row.PreScriptUninstall -and -not [string]::IsNullOrEmpty($row.PreScriptUninstall)) -or
    ($null -ne $row.PostScriptUninstall -and -not [string]::IsNullOrEmpty($row.PostScriptUninstall))
) {
    Write-Log "Checking Pre-Script and Post-Script Values for $($row.PackageID)"
    $row.PreScriptInstall
    $row.PostScriptInstall
    $row.PreScriptUninstall
    $row.PostScriptUninstall
    if ($row.PreScriptInstall -notlike "*.ps1" -or $row.PostScriptInstall -notlike "*.ps1" -or $row.PreScriptUninstall -notlike "*.ps1" -or $row.PostScriptUninstall -notlike "*.ps1" ) {
        Write-Log "Invalid post or pre-script for package $($row.PackageID) found in CSV. Check that the value contains: .ps1 - Please review the CSV. Exiting.." -ForegroundColor "Red"
        break
}
} 
#>

            #Print CustomArgumentListInstall if set and wait confirm
            Write-Log "Checking CustomArgumentListInstall Value for $($row.PackageID)"
            if ($row.CustomArgumentListInstall -ne '' -or $null) {
                Write-Log '-- CustomArgumentListInstall --'
                Write-Log "$($row.CustomArgumentListInstall)"
                if (!$SkipConfirmation) {
                    $confirmation = Read-Host "Please confirm CustomArgumentListInstall ($PackageID)? (Y/N)"

                    if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                        Write-Log 'Confirmed'
                    }
                    else {
                        Write-Log 'CustomArgumentListInstall not confirmed. Exiting..'
                        return
                    }     
                }
            }

            #Print CustomArgumentListUninstall if set and wait confirm
            Write-Log "Checking CustomArgumentListUninstall Value for $($row.PackageID)"
            if ($row.CustomArgumentListUninstall -ne '' -or $null) {
                Write-Log '-- CustomArgumentListUninstall --'
                Write-Log "$($row.CustomArgumentListUninstall)"
                if (!$SkipConfirmation) {
                    $confirmation = Read-Host "Please confirm CustomArgumentListUninstall ($PackageID)? (Y/N)"

                    if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
                        Write-Log 'Confirmed'
                    }
                    else {
                        Write-Log 'CustomArgumentListUninstall not confirmed. Exiting..'
                        return
                    }     
                }
            }

            Write-Log "Finished Validation for $($row.PackageID)"
            Write-Log '--- Validation End ---'
            Write-Log ''

            #Print Package details and wait for confirmation. If package not found break.
            $PackageIDOutLines = @(winget show --exact --id $($row.PackageID) --scope $($row.Context) --accept-source-agreements)
            #Check if targetversion specified
            if ($null -ne $row.TargetVersion -and $row.TargetVersion -ne '') {
                $PackageIDOutLines = @(winget show --exact --id $($row.PackageID) --scope $($row.Context) --version $($row.TargetVersion) --accept-source-agreements)
            }
            $PackageIDout = $PackageIDOutLines -join "`r`n"

            if ($PackageIDOutLines -notcontains 'No package found matching input criteria.') {
                if ($PackageIDOutLines -notcontains '  No applicable installer found; see logs for more details.') {
                    if (!$SkipConfirmation) {
                        Write-Log '--- WINGET PACKAGE INFORMATION ---'
                        Write-Log $PackageIDOut
                        Write-Log '--------------------------'
                        $confirmation = Read-Host 'Confirm the package details above (Y/N)'
                        if ($confirmation -eq 'N' -or $confirmation -eq 'N') {
                            break
                        }
                    }
                }
                else {
                    # Second condition not met
                    Write-Log "Applicable installer not found for $($row.Context) context" -ForegroundColor 'Red'
                    $imported = $False
                    $row | Add-Member -MemberType NoteProperty -Name 'Imported' -Value $imported  #Write imported status to $row
                    $errortext = "Applicable installer not found for $($row.Context) context"
                    $row | Add-Member -MemberType NoteProperty -Name 'ErrorText' -Value $errortext  #Write errortext to $row
                    continue
                }
            }
            else {
                Write-Log "Package $($row.PackageID) not found using winget" -ForegroundColor 'Red'
                return
            }


            #Scrape Winget package details to use in InTune from $PackageIDOut
            #Scrape "Found " Scrape all line from after this
            #Scrape "Description"  Regex.. Find "Description: " Scrape all line from after this

            #Clear variables
            $variables = 'PackageName', 'Version', 'Publisher', 'PublisherURL', 'PublisherSupportURL', 'Author', 'Description', 'Homepage', 'License', 'LicenseURL', 'Copyright', 'CopyrightURL', 'InstallerType', 'InstallerLocale', 'InstallerURL', 'InstallerSHA256'
            Remove-Variable -Name $variables -ErrorAction SilentlyContinue

            # Use regular expressions to extract details
            $PackageName = $PackageIDOut | Select-String -Pattern 'Found (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $PackageName = $PackageName -replace '\[', '(' -replace '\]', ')'  #Clean Packagename replace [] with ()
            $Version = $PackageIDOut | Select-String -Pattern 'Version: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $Publisher = $PackageIDOut | Select-String -Pattern 'Publisher: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $PublisherURL = $PackageIDOut | Select-String -Pattern 'Publisher Url: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $PublisherSupportURL = $PackageIDOut | Select-String -Pattern 'Publisher Support Url: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $Author = $PackageIDOut | Select-String -Pattern 'Author: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $Description = $PackageIDOut | Select-String -Pattern 'Description: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $Homepage = $PackageIDOut | Select-String -Pattern 'Homepage: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $License = $PackageIDOut | Select-String -Pattern 'License: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $LicenseURL = $PackageIDOut | Select-String -Pattern 'License Url: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $Copyright = $PackageIDOut | Select-String -Pattern 'Copyright: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $CopyrightURL = $PackageIDOut | Select-String -Pattern 'Copyright Url: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }

            # Extract Installer details
            $InstallerType = $PackageIDOut | Select-String -Pattern 'Installer Type: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $InstallerLocale = $PackageIDOut | Select-String -Pattern 'Installer Locale: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $InstallerURL = $PackageIDOut | Select-String -Pattern 'Installer Url: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }
            $InstallerSHA256 = $PackageIDOut | Select-String -Pattern 'Installer SHA256: (.+)' | ForEach-Object { $_.Matches.Groups[1].Value }

            # Loop through the variable names and assign "None" if empty or null  (Avoid script issues)
            foreach ($variable in $variables) {
                if ($null -eq (Get-Variable -Name $variable -ValueOnly)) {
                    Set-Variable -Name $variable -Value 'None'
                }
            }

            # Display the extracted variables
            Write-Log ''
            Write-Log "--- WinGet Scraped Metadata $($row.PackageID) ---"
            Write-Log "PackageName: $PackageName"
            Write-Log "Version: $Version"
            Write-Log "Publisher: $Publisher"
            Write-Log "PublisherURL: $PublisherURL"
            Write-Log "PublisherSupportURL: $PublisherSupportURL"
            Write-Log "Author: $Author"
            Write-Log "Description: $Description"
            Write-Log "Homepage: $Homepage"
            Write-Log "License: $License"
            Write-Log "LicenseURL: $LicenseURL"
            Write-Log "Copyright: $Copyright"
            Write-Log "CopyrightURL: $CopyrightURL"
            Write-Log "InstallerType: $InstallerType"
            Write-Log "InstallerLocale: $InstallerLocale"
            Write-Log "InstallerURL: $InstallerURL"
            Write-Log "InstallerSHA256: $InstallerSHA256"
            Write-Log ''
            Write-Log '--------------------------'

            # Build package details to prepare InTune import
            $PackageName += " ($($row.Context))"
            if ($Row.TargetVersion -ne $null) {
                $PackageName += " $($row.TargetVersion)"
            }

            $PackageName += ' (WinGet)'

            if ($Row.UpdateOnly -eq $true) {
                $PackageName = 'Update for ' + $PackageName
            }

            #Clear ArgumentList and CommandLine parameters.
            Remove-Variable -Name ArgumentListInstall, ArgumentListUninstall, InstallCommandLine, UninstallCommandLine -ErrorAction SilentlyContinue

            #Build ArgumentListInstall (if no custom argumentlist set)
            if ([string]::IsNullOrEmpty($row.CustomArgumentListInstall)) {
                $ArgumentListInstall = "install --exact --id $($row.PackageID) --silent --accept-package-agreements --accept-source-agreements --scope $($row.Context)"
                if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne '') {
                    $ArgumentListInstall += " --version $($Row.TargetVersion)"
                }
            }
            else {
                $ArgumentListInstall = $row.CustomArgumentListInstall
            }
            #Replace first 7 characters of ArgumentListInstall with "update" if package is UpdateOnly
            if ($Row.UpdateOnly -eq $true) {
                $ArgumentListInstall = 'update' + $ArgumentListInstall.Substring(7)
            }

            #Build ArgumentListUninstall (if no custom argumentlist set)
            if ([string]::IsNullOrEmpty($row.CustomArgumentListUninstall)) {
                $ArgumentListUninstall = "uninstall --exact --id $($row.PackageID) --silent --accept-source-agreements --scope $($row.Context)"
                if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne '') {
                    $ArgumentListUninstall += " --version $($Row.TargetVersion)"
                }
            }
            else {
                $ArgumentListUninstall = $row.CustomArgumentListUninstall
            }


            # Build install commandline
            $InstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName $($row.PackageID) -ArgumentList " + [char]34 + "$ArgumentListInstall" + [char]34
            if ($row.StopProcessInstall -ne $null -and $row.StopProcessInstall -ne '') {
                $InstallCommandLine += " -StopProcess '$($row.StopProcessInstall)'"
            }
            if ($row.PreScriptInstall -ne $null -and $row.PreScriptInstall -ne '') {
                $InstallCommandLine += " -Prescript '$($row.PreScriptInstall)'"
            }

            if ($row.PostScriptInstall -ne $null -and $row.PostScriptInstall -ne '') {
                $InstallCommandLine += " -Postscript '$($row.PostScriptInstall)'"
            }

            # Build uninstall commandline
            $UninstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName $($row.PackageID) -ArgumentList " + [char]34 + "$ArgumentListUninstall" + [char]34
            if ($row.StopProcessUninstall -ne $null -and $row.StopProcessUninstall -ne '') {
                $UninstallCommandLine += " -StopProcess '$($row.StopProcessUninstall)'"
            }

            if ($row.PreScriptUninstall -ne $null -and $row.PreScriptUninstall -ne '') {
                $InstallCommandLine += " -PreUninstall '$($row.PreScriptUninstall)'"
            }

            if ($row.PostScriptUninstall -ne $null -and $row.PostScriptUninstall -ne '') {
                $InstallCommandLine += " -PostUninstall '$($row.PostScriptUninstall)'"
            }

            Write-Log ' '
            Write-Log '--InstallCommandLine--'
            Write-Log "$InstallCommandLine"
            Write-Log ' '
            Write-Log '--UninstallCommandLine--'
            Write-Log "$UninstallCommandLine"
            Write-Log ' '

            #Make folder to store script files

            #Find script root - set $currentDirectory
            if ($scriptRoot) {
                $currentDirectory = $scriptRoot 
            }
            else {
                $currentDirectory = $PSScriptRoot
            }

            write-log "currentDirectory: $currentDirectory"
            write-log "scriptRoot: $scriptRoot"
            write-log "PSScriptRoot: $PSScriptRoot"
            write-log "csvFile: $csvFile"
            write-log "LogFile: $LogFile"
            write-log "TenantID: $TenantID"
            write-log "SkipConfirmation: $SkipConfirmation"
            write-log "SkipInTuneConnection: $SkipInTuneConnection"
            write-log "SkipModuleCheck: $SkipModuleCheck"
            # Log the retrieval of variables using Get-Variable
            $variables = Get-Variable
            foreach ($variable in $variables) {
                Write-Output "Retrieving variable: $($variable.Name) = $($variable.Value)"
            }

            # Generate foldername
            $folderName = "$($Row.PackageID)"
            $folderName += "-$($Row.Context)"

            if ($Row.UpdateOnly -eq $True) {
                $folderName += '-UpdateOnly'
            }

            if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne '') {
                $folderName += "-V$($Row.TargetVersion)"
            }

            $folderName += "-$currentUser-$Timestamp"

            # Combine the current directory and "packages" to create the full "Packages" directory path
            $packagesDirectory = Join-Path -Path $currentDirectory -ChildPath 'Packages'

            # Check if the "packages" directory exists, and if not, create it
            if (-Not (Test-Path -Path $packagesDirectory)) {
                New-Item -Path $packagesDirectory -ItemType Directory | Out-Null
            }

            # Combine the "Packages" directory path with the folder name to create the full path
            $PackageFolderPath = Join-Path -Path $packagesDirectory -ChildPath $folderName

            # Create the subfolder with the desired name
            New-Item -Path $PackageFolderPath -ItemType Directory | Out-Null

            Write-Log "Local Package Data Location: $PackageFolderPath"

            #Build DetectionScript
            $WingetDetectionScriptSource = "$currentDirectory\WinGet-WrapperDetection.ps1"
            $WingetDetectionScriptDestination = "$PackageFolderPath\WinGet-WrapperDetection-$($Row.PackageID).ps1"

            # Copy the source script to the destination
            Copy-Item -Path $WingetDetectionScriptSource -Destination $WingetDetectionScriptDestination

            # Read the destination script
            $scriptContent = Get-Content -Path $WingetDetectionScriptDestination

            # Identify the start and end of the # Settings section
            $startPattern = '# Settings'
            $endPattern = '# EndSettings'

            # Join the script content into a single string
            $scriptText = [string]::Join([Environment]::NewLine, $scriptContent)

            # Check if the # Settings section is present
            if ($scriptText -match "(?s)$startPattern.*?$endPattern") {
                $settingsSection = $Matches[0]  # Extract the # Settings section

                # If $Row.AcceptNewerVersion is $True, replace "$True" with "$True"
                if ($Row.AcceptNewerVersion -eq $True) {
                    $settingsSection = $settingsSection -replace '"\$True"', '$True'
                }
                # If $Row.AcceptNewerVersion is $False, replace "$True" with "$False"
                elseif ($Row.AcceptNewerVersion -eq $False) {
                    $settingsSection = $settingsSection -replace '"\$True"', '$False'
                }

                # Replace "Exact WinGet Package ID" with the actual value of $PackageID within the # Settings section
                $settingsSection = $settingsSection -replace 'Exact WinGet Package ID', $Row.PackageID

                # If $TargetVersion is not null or not empty, replace "TargetVersion = ''" with "TargetVersion = '$TargetVersion'" within the # Settings section
                if ($row.TargetVersion -ne $null -and $row.TargetVersion -ne '') {
                    $settingsSection = $settingsSection -replace 'TargetVersion = ""', "TargetVersion = ""$($Row.TargetVersion)"""
                }

                # Replace the updated # Settings section in the script content
                $scriptText = $scriptText -replace "(?s)$startPattern.*?$endPattern", $settingsSection

                # Save the updated script content
                Set-Content -Path $WingetDetectionScriptDestination -Value $scriptText

                Write-Log "Detection Script complete. $WingetDetectionScriptDestination"
            }
            else {
                Write-Log 'The # Settings section was not found in the script.' -ForegroundColor 'Red'
            }


            #Build RequirementScript
            if ($Row.UpdateOnly -eq $True) {
                $WingetRequirementScriptSource = "$currentDirectory\Winget-WrapperRequirements.ps1"
                $WingetRequirementScriptDestination = "$PackageFolderPath\WinGet-WrapperRequirements-$($Row.PackageID).ps1"
    
                # Copy the source script to the destination
                Copy-Item -Path $WingetRequirementScriptSource -Destination $WingetRequirementScriptDestination
    
                # Read the destination script
                $scriptContent = Get-Content -Path $WingetRequirementScriptDestination
    
                # Check if "Exact WinGet Package ID" is present in the script content
                if ($scriptContent -match 'Exact WinGet Package ID') {
                    # Replace "Exact WinGet Package ID" with the actual value of $PackageID
                    $scriptContent = $scriptContent -replace 'Exact WinGet Package ID', $Row.PackageID
    
            
                    # Save the updated script content
                    Set-Content -Path $WingetRequirementScriptDestination -Value $scriptContent
        
                    Write-Log "Requirement Script complete. $WingetRequirementScriptDestination"
                }
                else {
                    Write-Log "The text 'Exact WinGet Package ID' was not found in the script." -ForegroundColor 'Red'
                }
            }

            #Build Targetversion to use with intune package
            if ($row.TargetVersion -ne $null -and $row.TargetVersion -ne '') {
                $VersionInfo = $row.TargetVersion
            }
            else {
                $VersionInfo = 'Latest'
            }

            #Convert Context
            if ($row.Context -in @('Machine', 'machine')) {
                $ContextConverted = 'System'
            }
            elseif ($row.Context -in @('User', 'user')) {
                $ContextConverted = 'User'
            }

            #Build IntuneWin Folder  (including pre/postscripts if specified in CSV)
            # Create a folder named "InTuneWin" if it doesn't exist
            try {
                $intuneWinFolderPath = Join-Path -Path $PackageFolderPath -ChildPath 'InTuneWin'
                if (-not (Test-Path -Path $intuneWinFolderPath -PathType Container)) {
                    New-Item -Path $intuneWinFolderPath -ItemType Directory | Out-Null
                }
                # Copy WinGet-Wrapper.ps1
                Copy-Item -Path "$currentDirectory\WinGet-Wrapper.ps1" -Destination $intuneWinFolderPath -ErrorAction Stop

                # Copy pre or post script if specified
                if ($row.PreScriptInstall -ne $null -and $row.PreScriptInstall -ne '') {
                    Copy-Item -Path $row.PreScriptInstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PreScriptInstall) to $($intuneWinFolderPath)"
                }
                if ($row.PostScriptInstall -ne $null -and $row.PostScriptInstall -ne '') {
                    Copy-Item -Path $row.PostScriptInstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PostScriptInstall) to $($intuneWinFolderPath)"
                }
                if ($row.PreScriptUninstall -ne $null -and $row.PreScriptUninstall -ne '') {
                    Copy-Item -Path $row.PreScriptUninstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PreScriptUninstall) to $($intuneWinFolderPath)"
                }
                if ($row.PostScriptUninstall -ne $null -and $row.PostScriptUninstall -ne '') {
                    Copy-Item -Path $row.PostScriptUninstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PostScriptUninstall) to $($intuneWinFolderPath)"
                }
            }
            catch {
                Write-Log 'Failed to copy files to build InTuneWin package. Please check pre/post script is found'
                return
            }

            # Build WinGet-Wrapper.intunewin
            Write-Log 'Building WinGet-Wrapper.InTuneWin file'
            try {
                [string]$toolargs = "-c ""$($intuneWinFolderPath)"" -s ""WinGet-Wrapper.ps1"" -o ""$($PackageFolderPath)"" -q"
(Start-Process -FilePath "$currentDirectory\IntuneWinAppUtil.exe" -ArgumentList $toolargs -PassThru:$true -ErrorAction Stop -NoNewWindow).WaitForExit()
    
                # Check that IntuneWin was created
                if (Test-Path -Path "$PackageFolderPath\WinGet-Wrapper.intunewin" -PathType Leaf) {
                    Write-Log "IntuneWinAppUtil.exe built WinGet-Wrapper.intunewin for $PackageName"
                }
                else {
                    Write-Log 'IntuneWinAppUtil.exe could not build WinGet-Wrapper.intunewin'
                    return
                }
    
            }
            catch {
                Write-Log 'Error running IntuneWinAppUtil.exe'
                return
            }

            #Set WinGet-Wrapper.intunewin path
            $WinGetWrapperInTuneWinFilePath = Join-Path -Path $PackageFolderPath -ChildPath 'WinGet-Wrapper.intunewin'

            #Detection
            $DetectionRuleScript = New-IntuneWin32AppDetectionRuleScript -ScriptFile $WingetDetectionScriptDestination

            #RequirementRule Base
            $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture 'All' -MinimumSupportedWindowsRelease 'W10_20H2' 

            # Build base Add-InTuneWin32App Arguments
            $AddInTuneWin32AppArguments = @{
                FilePath             = "$WinGetWrapperInTuneWinFilePath"
                DisplayName          = "$PackageName"
                Description          = "$Description"
                Publisher            = "$Publisher"
                AppVersion           = "$VersionInfo"
                Developer            = "$Author"
                InstallCommandLine   = "$InstallCommandLine"
                UninstallCommandLine = "$UninstallCommandLine"
                InstallExperience    = "$ContextConverted"
                RestartBehavior      = 'suppress'
                Owner                = "$currentUser"
                Notes                = "Created by $currentUser at $Timestamp using WinGet-WrapperCreateFromCSV (https://github.com/SorenLundt/WinGet-Wrapper)"
                DetectionRule        = $DetectionRuleScript
                RequirementRule      = $RequirementRule
            }

            # Build AdditionalRequirementRule (if UpdateOnly True)
            if ($Row.UpdateOnly -eq $True) {
                $AdditionalRequirementRule = New-IntuneWin32AppRequirementRuleScript -ScriptFile $WingetRequirementScriptDestination -ScriptContext "$ContextConverted" -StringComparisonOperator 'equal' -StringOutputDataType -StringValue 'Installed'
                $AddInTuneWin32AppArguments.AdditionalRequirementRule = $AdditionalRequirementRule
            }
            else {
            }

            # If overwrite enabled , remove win32app from intune before importing
            # Disabled section. Needs more work.
            <#
if ($Overwrite -eq $True) {
    try{
    Write-Log "Overwrite enabled - Attempting to remove Win32App if it already exists in InTune - $PackageName "
    Write-Log "$packagename"
    $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName"
    Write-Log "627 - $CheckIntuneAppExists"
    if ($CheckIntuneAppExists.Count -gt 0)
    {
    Write-Log "Package found in InTune "
    Remove-IntuneWin32App -DisplayName "$PackageName"
    [int]$NumberOfTimesChecked = 1
    do {
        Write-Log "Checking if Win32App was removed before continuing. Checks: $NumberOfTimesChecked"
        Write-Log "635 - $CheckIntuneAppExists"
        $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName"
        Start-Sleep 15
        $NumberOfTimesChecked++
    } while ($CheckIntuneAppExists.Count -gt 0)
    }
}
    catch {
        Write-Log "An error occurred: $_.Exception.Message" -ForeGroundColor "Red"
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
        $errortext = "Error Remvoving Win32App: $_.Exception.Message"
        $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
        continue
    }

}
#>
    
            # Disabled section. Needs more work.
            <#  
Write-Log "Checking if application already exists in Intune - $PackageName"
    # Get the Intune Win32 apps with the specified display name
    $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName" -WarningAction SilentlyContinue

    # Check if any matching applications were found
    Write-Log "CheckInTuneAppExists.Count = $($CheckIntuneAppExists.Count)"
    if ($CheckIntuneAppExists.Count -gt 0) {
        Write-Log "ERROR: Application with the same name already exists in Intune. Could not import - $PackageName" -ForegroundColor "Red"
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
        $errortext = "Already exists in InTune"
        $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
        continue
    }
    else {
        Write-Log "OK! A similar package was not found in Intune."
    }
#>

            #Import application to InTune
            try {
                Write-Log "Importing application '$PackageName' to InTune"
                $AppImport = Add-IntuneWin32App @AddInTuneWin32AppArguments -WarningAction Stop
                $AppID = $AppImport.ID
                $row | Add-Member -MemberType NoteProperty -Name 'AppID' -Value $AppID #Write AppID to $row
                Write-Log "Imported application '$PackageName' to InTune" -ForeGroundColor 'Green'
                $imported = $True
            }
            catch {
                # Check if permission is missing and fix it
                if ($($_.Exception.Message) -match "DeviceManagementApps.ReadWrite.All") {
                    Connect-MgGraph -ClientId $ClientID -Scope 'DeviceManagementApps.ReadWrite.All' -TenantId $TenantID -ErrorAction Stop
                    Start-Sleep -Seconds 10
                    # Try again to upload after permission grant
                    try {
                        Write-Log "Importing application '$PackageName' to InTune"
                        $AppImport = Add-IntuneWin32App @AddInTuneWin32AppArguments -WarningAction Stop
                        $AppID = $AppImport.ID
                        $row | Add-Member -MemberType NoteProperty -Name 'AppID' -Value $AppID #Write AppID to $row
                        Write-Log "Imported application '$PackageName' to InTune" -ForeGroundColor 'Green'
                        $imported = $True
                    }
                    catch {
                        Write-Log "An error occurred: $($_.Exception.Message)" -ForeGroundColor 'Red'
                        $imported = $False
                        $row | Add-Member -MemberType NoteProperty -Name 'Imported' -Value $imported  #Write imported status to $row
                        $errortext = "Error Importing: $($_.Exception.Message)"
                        $row | Add-Member -MemberType NoteProperty -Name 'ErrorText' -Value $errortext  #Write errortext to $row
                        continue
                    }
                }
                else {
                    Write-Log "An error occurred: $($_.Exception.Message)" -ForeGroundColor 'Red'
                    $imported = $False
                    $row | Add-Member -MemberType NoteProperty -Name 'Imported' -Value $imported  #Write imported status to $row
                    $errortext = "Error Importing: $($_.Exception.Message)"
                    $row | Add-Member -MemberType NoteProperty -Name 'ErrorText' -Value $errortext  #Write errortext to $row
                    continue
                }

            }

            # Deploy Application if set - Install Intent "Required"
            try {
                if ($Row.InstallIntent -contains 'Required' -and ($null -ne $Row.GroupID -and $Row.GroupID -ne '')) {
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $Row.AppID -GroupID $Row.GroupID -Intent $row.InstallIntent -Notification $row.Notification -ErrorAction continue
                    Write-Log "Deployed AppID:$($Row.AppID) to $($Row.GroupID)" -ForeGroundColor 'Green'
                }
    
                # Deploy Application if set - Install Intent "Available"
                if ($Row.InstallIntent -contains 'Available' -and ($null -ne $Row.GroupID -and $Row.GroupID -ne '')) {
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $Row.AppID -GroupID $Row.GroupID -Intent $row.InstallIntent -Notification $row.Notification -ErrorAction continue
                    Write-Log "Deployed AppID:$($Row.AppID) to $($Row.GroupID)" -ForeGroundColor 'Green'
                }
            }
            catch {
                Write-Log "Error deploying $($Row.PackageID) : $_"
            }

            #Success
            $imported = $True
            $row | Add-Member -MemberType NoteProperty -Name 'Imported' -Value $imported  #Write imported status to $row

        }
        catch {
            #Failed
            $imported = $False
            $row | Add-Member -MemberType NoteProperty -Name 'Imported' -Value $imported  #Write imported status to $row
            $errortext = "Unknown Error: $_.Exception.Message"
            $row | Add-Member -MemberType NoteProperty -Name 'ErrorText' -Value $errortext  #Write errortext to $row

        }

        #If more packages to import, wait 15 seconds to avoid throttling. Microsoft Graph throttling ?
        # Check if there are more than 2 rows in $data
        if ($data.Count -gt 2) {
            Write-Log 'Waiting 15s before importing next package.. (Throttle)'
            Start-Sleep -Seconds 15
        }

    }
}
###### END FOREACH ######

#Write Results
Write-Log '---- RESULTS Package Creation ----'
foreach ($row in $data) {
    $importedStatus = $row.Imported
    $textColor = 'Green'  # Default to green
    $importedtext = 'Success' #Default to success
    if (-not $importedStatus) {
        $textColor = 'Red'  # Change to red if Imported is False
        $importedtext = 'Failed'
    }
    Write-Log ''
    $formattedText = "IMPORTED:$ImportedText PackageID:$($row.PackageID) AppID:$($row.AppID) TargetVersion:$($row.TargetVersion) Context:$($row.Context) UpdateOnly: $($row.UpdateOnly) AcceptNewerVersion: $($row.AcceptNewerVersion) ErrorText: $($row.ErrorText)"
    Write-Log $formattedText -ForegroundColor $textColor

    # If deployed also show these results
    if ($null = $row.GroupID -or $row.GroupID -ne '') {
        if ($importedStatus -eq $True) {
            Write-Log "> DEPLOYED:$($row.PackageID) AppID:$($row.AppID) ----> GroupID:$($row.GroupID) InstallIntent:$($row.InstallIntent) Notification:$($row.Notification)" -ForegroundColor $textColor
        }   
    }
    Write-Host "Log File: $LogFile"
}
