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
# Version 3.0 - 16-07-2025 - Replaced IntuneWin32App module with native Microsoft.Graph.Intune cmdlets

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
    [Switch]$SkipModuleCheck = $false,

    #Skips Graph Connection if already connected from a parent process
    [Switch]$SkipGraphConnection = $false
)

#Timestamp
$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Default Log File
if (-not $LogFile) {
    $LogFolder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"

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

    $LogMsgTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$LogMsgTimeStamp] $Message"

    # Output to the console
    Write-Host $LogEntry

    # Append to the log file
    $LogEntry | Out-File -Append -FilePath $LogFile
}
Write-Log "Log File: $LogFile"

Write-log "Parameters: csvFile: $csvFile, TenantID: $TenantID, SkipConfirmation: $SkipConfirmation, LogFile: $LogFile, SkipModuleCheck: $SkipModuleCheck, scriptRoot: $scriptRoot"

# Install and load required modules
if (-not $SkipModuleCheck) {
    $requiredModules = @{
        "IntuneWin32App" = "DeviceManagementApps.ReadWrite.All"
        "Microsoft.Graph.Authentication" = "User.Read"
    }

    foreach ($moduleName in $requiredModules.Keys) {
        $moduleInstalled = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue
        if (-not $moduleInstalled) {
            Write-Log "Installing module $moduleName..."
            Install-Module -Name $moduleName -Force -AllowClobber -Scope CurrentUser
        }
        else {
            $latestVersion = (Find-Module -Name $moduleName).Version
            if ($moduleInstalled.Version -lt $latestVersion) {
                Write-Log "Updating module $moduleName..."
                Update-Module -Name $moduleName -Force
            }
            else {
                Write-Log "Module $moduleName is already up-to-date."
            }
        }
    }
}
#Import modules
Import-Module -Name "IntuneWin32App"
Import-Module -Name "Microsoft.Graph.Authentication"

# Welcome greeting
Write-Log " "
Write-Log " "
Write-Log "-----------------------------"
Write-Log "---- WinGet-WrapperCreate----"
Write-Log "-----------------------------"
Write-Log " "
Write-Log " "
Write-Log "https://github.com/SorenLundt/WinGet-Wrapper"
Write-Log "       GNU General Public License v3"

# Test CSV path
if (Test-Path -Path "$csvFile" -PathType Leaf) {
    Write-Log "File: $csvFile"
}
else {
    Write-Log "File not found: $csvFile" -ForegroundColor "Red"
    return
}

# Import the CSV file with custom headers
$data = Import-Csv -Path "$csvFile" -Header "PackageID", "Context", "AcceptNewerVersion", "UpdateOnly", "TargetVersion", "StopProcessInstall", "StopProcessUninstall", "PreScriptInstall", "PostScriptInstall", "PreScriptUninstall", "PostScriptUninstall", "CustomArgumentListInstall", "CustomArgumentListUninstall", "InstallIntent", "Notification", "GroupID" | Select-Object -Skip 1
Write-Log "Found $($data.Count) packages to process in $csvFile"

# Convert "AcceptNewerVersion" and "UpdateOnly" columns to Boolean values
$data = $data | ForEach-Object {
    $_.AcceptNewerVersion = [bool]($_.AcceptNewerVersion -as [int])
    $_.UpdateOnly = [bool]($_.UpdateOnly -as [int])
    $_
}
Write-Log "-- IMPORT LIST --"
foreach ($row in $data) {
    Write-Log "IMPORT PackageID:$($row.PackageID) - Context:$($row.Context) - UpdateOnly:$($row.UpdateOnly) - TargetVersion:$($row.TargetVersion)" -ForegroundColor Gray
}
Write-Log ""

Write-Log "-- DEPLOY LIST --"
foreach ($row in $data) {
    if ($null = $row.GroupID -or $row.GroupID -ne "") {
        Write-Log "DEPLOY PackageID:$($row.PackageID) GroupID:$($row.GroupID) InstallIntent:$($row.InstallIntent) Notification:$($row.Notification)" -ForegroundColor Gray
    }
}

#Connect to Intune
if (-not $SkipGraphConnection) {
    try {
        Write-Log "Connecting to Microsoft Graph..."
        Connect-MSIntuneGraph -TenantID $TenantID -ClientID $ClientID -RedirectUri $RedirectURL -Interactive
        Write-Log "Successfully connected to Microsoft Graph"
    }
    catch {
        Write-Log "ERROR: Connect-MSIntuneGraph Failed. Exiting. Details: $($_.Exception.Message)" -ForegroundColor "Red"
        break
    }
}
else {
    Write-Log "Skipping Graph connection as per -SkipGraphConnection parameter. Assuming existing connection."
}


#Import each application to InTune
foreach ($row in $data) {
    & {
        try {
            #TimeStamp
            $Timestamp = (Get-Date).ToString("yyyyMMddHHmmss")

            #Get User
            $currentUser = $env:USERNAME

            Write-Log "=================================================="
            Write-Log "Starting processing for Package: $($row.PackageID)"
            Write-Log "=================================================="

            Write-Log "--- Validation Start ---"
            #Check context is valid
            if ($row.Context -notin @("Machine", "machine", "User", "user", "None", "")) {
                Write-Log "FAILED: Invalid context setting '$($row.Context)'. Must be 'Machine', 'User', or 'None'." -ForegroundColor "Red"
                break
            }
            Write-Log "OK: Context is '$($row.Context)'"

            #Check AcceptNewerVersion is true or false
            if ($row.AcceptNewerVersion -ne $True -and $row.AcceptNewerVersion -ne $False) {
                Write-Log "FAILED: Invalid AcceptNewerVersion setting '$($row.AcceptNewerVersion)'. Must be '0' or '1'." -ForegroundColor "Red"
                break
            }
            Write-Log "OK: AcceptNewerVersion is '$($row.AcceptNewerVersion)'"

            #Check InstallIntent is valid
            if ($null -ne $row.InstallIntent -and $row.InstallIntent -ne "") {
                if ($row.InstallIntent -notcontains "Required" -and $row.InstallIntent -notcontains "required" -and $row.InstallIntent -notcontains "Available" -and $row.InstallIntent -notcontains "available") {
                    Write-Log "FAILED: Invalid InstallIntent setting '$($row.InstallIntent)'. Use 'Available' or 'Required'." -ForegroundColor "Red"
                    break
                }
                Write-Log "OK: InstallIntent is '$($row.InstallIntent)'"
            }

            #Check Notification is valid
            if ($null -ne $row.Notification -and $row.Notification -ne "") {
                $validNotificationValues = "showAll", "showReboot", "hideAll"
                if ($validNotificationValues -notcontains $row.Notification.ToLower()) {
                    Write-Log "FAILED: Invalid Notification setting '$($row.Notification)'. Use 'showAll', 'showReboot', or 'hideAll'." -ForegroundColor "Red"
                    break
                }
                Write-Log "OK: Notification is '$($row.Notification)'"
            }

            #Check GroupID is set if InstallIntent is specified
            if ($null -ne $row.InstallIntent -and $row.InstallIntent -ne "") {
                if ($row.GroupID -eq $null -or $row.GroupID -eq "") {
                    Write-Log "FAILED: GroupID is required when InstallIntent is set." -ForegroundColor "Red"
                    break
                }
                Write-Log "OK: GroupID is '$($row.GroupID)'"
            }

            #Check UpdateOnly is true or false
            if ($row.UpdateOnly -ne $True -and $row.UpdateOnly -ne $False) {
                Write-Log "FAILED: Invalid UpdateOnly setting '$($row.UpdateOnly)'. Must be '0' or '1'." -ForegroundColor "Red"
                break
            }
            Write-Log "OK: UpdateOnly is '$($row.UpdateOnly)'"

            #Check StopProcessInstall and StopProcessUninstall does not contain "exe"  (inform should not contain .exe)
            if ($row.StopProcessInstall -contains ".") {
                Write-Log "FAILED: Invalid StopProcessInstall setting '$($row.StopProcessInstall)'. Do not include file extensions." -ForegroundColor "Red"
                break
            }
            Write-Log "OK: StopProcessInstall check passed."
            if ($row.StopProcessUninstall -contains ".") {
                Write-Log "FAILED: Invalid StopProcessUninstall setting '$($row.StopProcessUninstall)'. Do not include file extensions." -ForegroundColor "Red"
                break
            }
            Write-Log "OK: StopProcessUninstall check passed."

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
            if ($row.CustomArgumentListInstall -ne "" -or $null) {
                Write-Log "INFO: CustomArgumentListInstall is set to '$($row.CustomArgumentListInstall)'"
                if (!$SkipConfirmation) {
                    $confirmation = Read-Host "Please confirm CustomArgumentListInstall for $($row.PackageID)? (Y/N)"

                    if ($confirmation -eq "Y" -or $confirmation -eq "y") {
                        Write-Log "OK: Confirmed CustomArgumentListInstall."
                    }
                    else {
                        Write-Log "ABORT: CustomArgumentListInstall not confirmed by user. Exiting."
                        return
                    }     
                }
            }

            #Print CustomArgumentListUninstall if set and wait confirm
            if ($row.CustomArgumentListUninstall -ne "" -or $null) {
                Write-Log "INFO: CustomArgumentListUninstall is set to '$($row.CustomArgumentListUninstall)'"
                if (!$SkipConfirmation) {
                    $confirmation = Read-Host "Please confirm CustomArgumentListUninstall for $($row.PackageID)? (Y/N)"

                    if ($confirmation -eq "Y" -or $confirmation -eq "y") {
                        Write-Log "OK: Confirmed CustomArgumentListUninstall."
                    }
                    else {
                        Write-Log "ABORT: CustomArgumentListUninstall not confirmed by user. Exiting."
                        return
                    }     
                }
            }

            Write-Log "--- Validation End ---"
            Write-Log ""

            #Print Package details and wait for confirmation.
            $wingetArgs = @(
                "show",
                "--exact",
                "--id", $row.PackageID,
                "--accept-source-agreements"
            )
            if ($row.Context -and $row.Context -ne "None") {
                $wingetArgs += @("--scope", $row.Context)
            }
            if ($null -ne $row.TargetVersion -and $row.TargetVersion -ne "") {
                $wingetArgs += @("--version", $row.TargetVersion)
            }
            Write-Log "Executing: winget $($wingetArgs -join ' ')"
            $PackageIDOutLines = & winget $wingetArgs
            $PackageIDout = $PackageIDOutLines -join "`r`n"

            if (!$SkipConfirmation) {
                Write-Log "--- WINGET PACKAGE INFORMATION ---"
                Write-Log $PackageIDOut
                Write-Log "--------------------------"
                $confirmation = Read-Host "Confirm the package details above (Y/N)"
                if ($confirmation -eq "N" -or $confirmation -eq "n") {
                    Write-Log "ABORT: User did not confirm package details. Skipping to next package."
                    continue
                }
            }

            #Scrape Winget package details to use in InTune from $PackageIDOut
            Write-Log "Scraping metadata from WinGet output..."
            #Scrape "Found " Scrape all line from after this
            #Scrape "Description"  Regex.. Find "Description: " Scrape all line from after this

            #Clear variables
            $variables = "PackageName", "Version", "Publisher", "PublisherURL", "PublisherSupportURL", "Author", "Description", "Homepage", "License", "LicenseURL", "Copyright", "CopyrightURL", "InstallerType", "InstallerLocale", "InstallerURL", "InstallerSHA256"
            Remove-Variable -Name $variables -ErrorAction SilentlyContinue

            # Use regular expressions to extract details
            $PackageName = $PackageIDOut | Select-String -Pattern "Found (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $PackageName = $PackageName -replace "\[", "(" -replace "\]", ")"  #Clean Packagename replace [] with ()
            $Version = $PackageIDOut | Select-String -Pattern "Version: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $Publisher = $PackageIDOut | Select-String -Pattern "Publisher: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $PublisherURL = $PackageIDOut | Select-String -Pattern "Publisher Url: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $PublisherSupportURL = $PackageIDOut | Select-String -Pattern "Publisher Support Url: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $Author = $PackageIDOut | Select-String -Pattern "Author: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $Description = $PackageIDOut | Select-String -Pattern "Description: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $Homepage = $PackageIDOut | Select-String -Pattern "Homepage: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $License = $PackageIDOut | Select-String -Pattern "License: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $LicenseURL = $PackageIDOut | Select-String -Pattern "License Url: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $Copyright = $PackageIDOut | Select-String -Pattern "Copyright: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $CopyrightURL = $PackageIDOut | Select-String -Pattern "Copyright Url: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }

            # Extract Installer details
            $InstallerType = $PackageIDOut | Select-String -Pattern "Installer Type: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $InstallerLocale = $PackageIDOut | Select-String -Pattern "Installer Locale: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $InstallerURL = $PackageIDOut | Select-String -Pattern "Installer Url: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            $InstallerSHA256 = $PackageIDOut | Select-String -Pattern "Installer SHA256: (.+)" | ForEach-Object { $_.Matches.Groups[1].Value }

            # Loop through the variable names and assign "None" if empty or null  (Avoid script issues)
            foreach ($variable in $variables) {
                if ($null -eq (Get-Variable -Name $variable -ValueOnly)) {
                    Set-Variable -Name $variable -Value "None"
                }
            }

            # Display the extracted variables
            Write-Log ""
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
            Write-Log ""
            Write-Log "--------------------------"

            # Build package details to prepare InTune import
            $OriginalPackageName = $PackageName
            $PackageName += " ($($row.Context))"
            if ($Row.TargetVersion -ne $null) {
                $PackageName += " $($row.TargetVersion)"
            }

            $PackageName += " (WinGet)"

            if ($Row.UpdateOnly -eq $true) {
                $PackageName = "Update for " + $PackageName
            }

            #Clear ArgumentList and CommandLine parameters.
            Remove-Variable -Name ArgumentListInstall, ArgumentListUninstall, InstallCommandLine, UninstallCommandLine -ErrorAction SilentlyContinue

            #Build ArgumentListInstall (if no custom argumentlist set)
            if ([string]::IsNullOrEmpty($row.CustomArgumentListInstall)) {
                $ArgumentListInstall = "install --exact --id $($row.PackageID) --silent --accept-package-agreements --accept-source-agreements"
                if ($row.Context -and $row.Context -ne "None") {
                    $ArgumentListInstall += " --scope $($row.Context)"
                }
                if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne "") {
                    $ArgumentListInstall += " --version $($Row.TargetVersion)"
                }
            }
            else {
                $ArgumentListInstall = $row.CustomArgumentListInstall
            }
            #Replace first 7 characters of ArgumentListInstall with "update" if package is UpdateOnly
            if ($Row.UpdateOnly -eq $true) {
                $ArgumentListInstall = "update" + $ArgumentListInstall.Substring(7)
            }

            #Build ArgumentListUninstall (if no custom argumentlist set)
            if ([string]::IsNullOrEmpty($row.CustomArgumentListUninstall)) {
                $ArgumentListUninstall = "uninstall --exact --id $($row.PackageID) --silent --accept-source-agreements"
                if ($row.Context -and $row.Context -ne "None") {
                    $ArgumentListUninstall += " --scope $($row.Context)"
                }
                if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne "") {
                    $ArgumentListUninstall += " --version $($Row.TargetVersion)"
                }
            }
            else {
                $ArgumentListUninstall = $row.CustomArgumentListUninstall
            }

            # Build install commandline with granular parameters
            $InstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageID `"$($row.PackageID)`""
            if ($row.Context -and $row.Context -ne "None") {
                $InstallCommandLine += " -Scope `"$($row.Context)`""
            }
            
            # Add UpdateOnly switch if specified
            if ($row.UpdateOnly -eq "1" -or $row.UpdateOnly -eq $true) {
                $InstallCommandLine += " -UpdateOnly"
            }
            
            # Add TargetVersion if specified
            if (![string]::IsNullOrEmpty($row.TargetVersion)) {
                $InstallCommandLine += " -TargetVersion `"$($row.TargetVersion)`""
            }
            
            # Add custom arguments if specified
            if (![string]::IsNullOrEmpty($row.CustomArgumentListInstall)) {
                $InstallCommandLine += " -CustomArgs `"$($row.CustomArgumentListInstall)`""
            }
            
            # Add process stopping if specified
            if (![string]::IsNullOrEmpty($row.StopProcessInstall)) {
                $InstallCommandLine += " -StopProcess `"$($row.StopProcessInstall)`""
            }
            
            # Add pre-script if specified
            if (![string]::IsNullOrEmpty($row.PreScriptInstall)) {
                $InstallCommandLine += " -PreScript `"$($row.PreScriptInstall)`""
            }
            
            # Add post-script if specified
            if (![string]::IsNullOrEmpty($row.PostScriptInstall)) {
                $InstallCommandLine += " -PostScript `"$($row.PostScriptInstall)`""
            }

            # Build uninstall commandline with granular parameters
            $UninstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageID `"$($row.PackageID)`" -Uninstall"
            if ($row.Context -and $row.Context -ne "None") {
                $UninstallCommandLine += " -Scope `"$($row.Context)`""
            }
            
            # Add TargetVersion for uninstall if specified
            if (![string]::IsNullOrEmpty($row.TargetVersion)) {
                $UninstallCommandLine += " -TargetVersion `"$($row.TargetVersion)`""
            }
            
            # Add custom arguments for uninstall if specified
            if (![string]::IsNullOrEmpty($row.CustomArgumentListUninstall)) {
                $UninstallCommandLine += " -CustomArgs `"$($row.CustomArgumentListUninstall)`""
            }
            
            # Add process stopping for uninstall if specified
            if (![string]::IsNullOrEmpty($row.StopProcessUninstall)) {
                $UninstallCommandLine += " -StopProcess `"$($row.StopProcessUninstall)`""
            }
            
            # Add pre-script for uninstall if specified
            if (![string]::IsNullOrEmpty($row.PreScriptUninstall)) {
                $UninstallCommandLine += " -PreScript `"$($row.PreScriptUninstall)`""
            }
            
            # Add post-script for uninstall if specified
            if (![string]::IsNullOrEmpty($row.PostScriptUninstall)) {
                $UninstallCommandLine += " -PostScript `"$($row.PostScriptUninstall)`""
            }

            Write-Log ' '
            Write-Log '--InstallCommandLine--'
            Write-Log $InstallCommandLine
            Write-Log '--UninstallCommandLine--'
            Write-Log $UninstallCommandLine
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
                $folderName += "-UpdateOnly"
            }

            if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne "") {
                $folderName += "-V$($Row.TargetVersion)"
            }

            $folderName += "-$currentUser-$Timestamp"

            # Combine the current directory and "packages" to create the full "Packages" directory path
            $packagesDirectory = Join-Path -Path $currentDirectory -ChildPath "Packages"

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
            $startPattern = "# Settings"
            $endPattern = "# EndSettings"

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
                $settingsSection = $settingsSection -replace "Exact WinGet Package ID", $Row.PackageID

                # If $TargetVersion is not null or not empty, replace "TargetVersion = ''" with "TargetVersion = '$TargetVersion'" within the # Settings section
                if ($row.TargetVersion -ne $null -and $row.TargetVersion -ne "") {
                    $settingsSection = $settingsSection -replace 'TargetVersion = ""', "TargetVersion = ""$($Row.TargetVersion)"""
                }

                # Replace the updated # Settings section in the script content
                $scriptText = $scriptText -replace "(?s)$startPattern.*?$endPattern", $settingsSection

                # Save the updated script content
                Set-Content -Path $WingetDetectionScriptDestination -Value $scriptText

                Write-Log "Detection Script complete. $WingetDetectionScriptDestination"
            }
            else {
                Write-Log "The # Settings section was not found in the script." -ForegroundColor "Red"
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
                if ($scriptContent -match "Exact WinGet Package ID") {
                    # Replace "Exact WinGet Package ID" with the actual value of $PackageID
                    $scriptContent = $scriptContent -replace "Exact WinGet Package ID", $Row.PackageID
    
            
                    # Save the updated script content
                    Set-Content -Path $WingetRequirementScriptDestination -Value $scriptContent
        
                    Write-Log "Requirement Script complete. $WingetRequirementScriptDestination"
                }
                else {
                    Write-Log "The text 'Exact WinGet Package ID' was not found in the script." -ForegroundColor "Red"
                }
            }

            #Build Targetversion to use with intune package
            if ($row.TargetVersion -ne $null -and $row.TargetVersion -ne "") {
                $VersionInfo = $row.TargetVersion
            }
            else {
                $VersionInfo = "Latest"
            }

            #Convert Context to proper Graph API values
            if ($row.Context -in @("Machine", "machine")) {
                $runAsAccount = "system"
            }
            elseif ($row.Context -in @("User", "user")) {
                $runAsAccount = "user"
            }
            else { # Default for "None" or empty
                $runAsAccount = "system"
            }

            #Build IntuneWin Folder  (including pre/postscripts if specified in CSV)
            # Create a folder named "InTuneWin" if it doesn't exist
            try {
                $intuneWinFolderPath = Join-Path -Path $PackageFolderPath -ChildPath "InTuneWin"
                if (-not (Test-Path -Path $intuneWinFolderPath -PathType Container)) {
                    New-Item -Path $intuneWinFolderPath -ItemType Directory | Out-Null
                }
                # Copy WinGet-Wrapper.ps1
                Copy-Item -Path "$currentDirectory\WinGet-Wrapper.ps1" -Destination $intuneWinFolderPath -ErrorAction Stop

                # Copy pre or post script if specified
                if ($row.PreScriptInstall -ne $null -and $row.PreScriptInstall -ne "") {
                    Copy-Item -Path $row.PreScriptInstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PreScriptInstall) to $($intuneWinFolderPath)"
                }
                if ($row.PostScriptInstall -ne $null -and $row.PostScriptInstall -ne "") {
                    Copy-Item -Path $row.PostScriptInstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PostScriptInstall) to $($intuneWinFolderPath)"
                }
                if ($row.PreScriptUninstall -ne $null -and $row.PreScriptUninstall -ne "") {
                    Copy-Item -Path $row.PreScriptUninstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PreScriptUninstall) to $($intuneWinFolderPath)"
                }
                if ($row.PostScriptUninstall -ne $null -and $row.PostScriptUninstall -ne "") {
                    Copy-Item -Path $row.PostScriptUninstall -Destination $intuneWinFolderPath -ErrorAction Stop
                    Write-Log "Copied $($row.PostScriptUninstall) to $($intuneWinFolderPath)"
                }
            }
            catch {
                Write-Log "Failed to copy files to build InTuneWin package. Please check pre/post script is found"
                return
            }

            # Build WinGet-Wrapper.intunewin
            Write-Log "Building WinGet-Wrapper.InTuneWin file..."
            try {
                [string]$toolargs = "-c ""$($intuneWinFolderPath)"" -s ""WinGet-Wrapper.ps1"" -o ""$($PackageFolderPath)"" -q"
                Write-Log "Running IntuneWinAppUtil.exe with args: $toolargs"
                (Start-Process -FilePath "$currentDirectory\IntuneWinAppUtil.exe" -ArgumentList $toolargs -PassThru:$true -ErrorAction Stop -NoNewWindow).WaitForExit()
    
                # Check that IntuneWin was created
                if (Test-Path -Path "$PackageFolderPath\WinGet-Wrapper.intunewin" -PathType Leaf) {
                    Write-Log "SUCCESS: IntuneWinAppUtil.exe built WinGet-Wrapper.intunewin for $PackageName"
                }
                else {
                    Write-Log "FAILED: IntuneWinAppUtil.exe could not build WinGet-Wrapper.intunewin"
                    return
                }
    
            }
            catch {
                Write-Log "FAILED: Error running IntuneWinAppUtil.exe. Details: $($_.Exception.Message)"
                return
            }

            #Set WinGet-Wrapper.intunewin path
            $WinGetWrapperInTuneWinFilePath = Join-Path -Path $PackageFolderPath -ChildPath "WinGet-Wrapper.intunewin"

            # --- Start IntuneWin32App Creation ---

            # 1. Detection Rule
            $DetectionRuleScript = New-IntuneWin32AppDetectionRuleScript -ScriptFile $WingetDetectionScriptDestination

            # 2. Requirement Rules
            $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease "W10_21H2"
            
            # Build base Add-IntuneWin32App Arguments
            $AddIntuneWin32AppArguments = @{
                FilePath             = "$WinGetWrapperInTuneWinFilePath"
                DisplayName          = "$PackageName"
                Description          = "$Description"
                Publisher            = "$Publisher"
                AppVersion           = "$VersionInfo"
                Developer            = "$Author"
                InstallCommandLine   = "$InstallCommandLine"
                UninstallCommandLine = "$UninstallCommandLine"
                InstallExperience    = "$runAsAccount"
                RestartBehavior      = "suppress"
                Owner                = "$currentUser"
                Notes                = "Created by $currentUser at $Timestamp using WinGet-WrapperCreateFromCSV (https://github.com/SorenLundt/WinGet-Wrapper)"
                DetectionRule        = $DetectionRuleScript
                RequirementRule      = $RequirementRule
            }

            # Build AdditionalRequirementRule (if UpdateOnly True)
            if ($Row.UpdateOnly -eq $True) {
                $AdditionalRequirementRule = New-IntuneWin32AppRequirementRuleScript -ScriptFile $WingetRequirementScriptDestination -ScriptContext $runAsAccount -StringComparisonOperator "equal" -StringOutputDataType -StringValue "Installed"
                $AddIntuneWin32AppArguments.AdditionalRequirementRule = $AdditionalRequirementRule
            }

            try {
                Write-Log "Importing application '$PackageName' to Intune..."
                $AppImport = Add-IntuneWin32App @AddIntuneWin32AppArguments -WarningAction Stop
                $AppID = $AppImport.ID
                $row | Add-Member -MemberType NoteProperty -Name "AppID" -Value $AppID #Write AppID to $row
                Write-Log "Successfully imported application '$PackageName' to Intune" -ForeGroundColor "Green"
                $imported = $True
            }
            catch {
                Write-Log "An error occurred during Intune import: $($_.Exception.Message)" -ForeGroundColor "Red"
                $imported = $False
                $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
                $errortext = "Error Importing: $($_.Exception.Message)"
                $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
                continue
            }

            # Deploy Application if set
            try {
                if ($imported -and $Row.InstallIntent -and $Row.GroupID) {
                    Write-Log "Attempting to create assignment for AppID: $($AppID) to GroupID: $($Row.GroupID)"
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $AppID -GroupID $Row.GroupID -Intent $row.InstallIntent -Notification $row.Notification -ErrorAction continue
                    Write-Log "SUCCESS: Deployed AppID:$($AppID) to $($Row.GroupID) with intent '$($Row.InstallIntent)'" -ForeGroundColor "Green"
                }
            }
            catch {
                Write-Log "Error deploying $($Row.PackageID) : $($_.Exception.Message)"
            }

            #Success
            $imported = $True
            $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
            Write-Log "Finished processing for Package: $($row.PackageID)"
        }
        catch {
            #Failed
            $imported = $False
            $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
            $errortext = "Unknown Error: $_.Exception.Message"
            $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row

        }

        #If more packages to import, wait 15 seconds to avoid throttling. Microsoft Graph throttling ?
        # Check if there are more than 2 rows in $data
        if ($data.Count -gt 2) {
            Write-Log "Waiting 15s before importing next package.. (Throttle)"
            Start-Sleep -Seconds 15
        }

    }
}
###### END FOREACH ######

#Write Results
Write-Log "---- RESULTS Package Creation ----"
foreach ($row in $data) {
    $importedStatus = $row.Imported
    $textColor = "Green"  # Default to green
    $importedtext = "Success" #Default to success
    if (-not $importedStatus) {
        $textColor = "Red"  # Change to red if Imported is False
        $importedtext = "Failed"
    }
    Write-Log ""
    $formattedText = "IMPORTED:$ImportedText PackageID:$($row.PackageID) AppID:$($row.AppID) TargetVersion:$($row.TargetVersion) Context:$($row.Context) UpdateOnly: $($row.UpdateOnly) AcceptNewerVersion: $($row.AcceptNewerVersion) ErrorText: $($row.ErrorText)"
    Write-Log $formattedText -ForegroundColor $textColor

    # If deployed also show these results
    if ($null = $row.GroupID -or $row.GroupID -ne "") {
        if ($importedStatus -eq $True) {
            Write-Log "> DEPLOYED:$($row.PackageID) AppID:$($row.AppID) ----> GroupID:$($row.GroupID) InstallIntent:$($row.InstallIntent) Notification:$($row.Notification)" -ForegroundColor $textColor
        }   
    }
    Write-Host "Log File: $LogFile"
}
