# Soren Lundt - 12-09-2023 - https://github.com/SorenLundt/WinGet-Wrapper
# Imports packages from WinGet to InTune (incuding available WinGet package metadata)
# Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm
# 
# Usage: .\WinGet-WrapperImportFromCSV.ps1 -TenantID company.onmicrosoft.com -csvFile WinGet-WrapperImportFromCSV.csv -SkipConfirmation
#
# Parameters:
# csvFile = csvFile to import from (default: WinGet-WrapperImportFromCSV.csv)
# TenantID = TenantID to connect to MSGraph/InTune
# SkipConfirmation = Skips confirmation for each package
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

#Parameters
Param (
    #CSV File to import from (default: WinGet-WrapperImportFromCSV.csv)
    [Parameter()]
    [string]$csvFile = "WinGet-WrapperImportFromCSV.csv",

    #TenantID to connect to MSGraph/InTune
    [Parameter(Mandatory=$True)]
    [string]$TenantID,

    #Skips confirmation for each package before import
    [Switch]$SkipConfirmation = $false
)

# Install and load required modules
Install-Module -Name "IntuneWin32App"  # https://github.com/MSEndpointMgr/IntuneWin32App
Import-Module -Name "IntuneWin32App"

Install-Module -Name "Microsoft.Graph.Intune"
Import-Module -Name "Microsoft.Graph.Intune"


# Welcome greeting
Write-host " "
Write-host " "
Write-host "-----------------------------"
Write-Host "---- WinGet-WrapperCreate----"
Write-host "-----------------------------"
write-host " "
write-host " "

# Test CSV path
if (Test-Path -Path $csvFile -PathType Leaf) {
    Write-Host "File: $csvFile"
} else {
    Write-Host "File not found: $csvFile" -ForegroundColor "Red"
    return
}

# Import the CSV file with custom headers
$data = Import-Csv -Path $csvFile -Header "PackageID", "Context", "AcceptNewerVersion", "UpdateOnly", "TargetVersion", "StopProcessInstall", "StopProcessUninstall", "PreScriptInstall", "PostScriptInstall", "PreScriptUninstall", "PostScriptUninstall", "CustomArgumentListInstall", "CustomArgumentListUninstall", "InstallIntent", "Notification", "GroupID" | Select-Object -Skip 1

# Convert "AcceptNewerVersion" and "UpdateOnly" columns to Boolean values
$data = $data | ForEach-Object {
    $_.AcceptNewerVersion = [bool]($_.AcceptNewerVersion -as [int])
    $_.UpdateOnly = [bool]($_.UpdateOnly -as [int])
    $_
}
Write-host "-- IMPORT LIST --"
foreach ($row in $data){
    Write-Host "IMPORT PackageID:$($row.PackageID) - Context:$($row.Context) - UpdateOnly:$($row.UpdateOnly) - TargetVersion:$($row.TargetVersion)" -ForegroundColor Gray
}
write-host ""

Write-host "-- DEPLOY LIST --"
foreach ($row in $data){
    if ($null = $row.GroupID -or $row.GroupID -ne "")
    {
        write-host "DEPLOY PackageID:$($row.PackageID) GroupID:$($row.GroupID) InstallIntent:$($row.InstallIntent) Notification:$($row.Notification)" -ForegroundColor Gray
    }
}

#Connect to Intune
try{  
Connect-MSIntuneGraph -TenantID "$TenantID" -Interactive
}
catch {
    write-host "ERROR: Connect-MSIntuneGraph Failed. Exiting" -ForegroundColor "Red"
    break
}

#Import each application to InTune
foreach ($row in $data) {
& {
    try{
#Write-Host "--- Package Details ---"
#Write-Host "PackageID: $($row.PackageID)"
#Write-Host "Context: $($row.Context)"
#Write-Host "AcceptNewerVersion: $($row.AcceptNewerVersion)"
#Write-Host "UpdateOnly: $($row.UpdateOnly)"
#Write-Host "TargetVersion: $($row.TargetVersion)"
#Write-Host "StopProcessInstall: $($row.StopProcessInstall)"
#Write-Host "StopProcessUninstall: $($row.StopProcessUninstall)"
#Write-Host "PreScriptInstall: $($row.PreScriptInstall)"
#Write-Host "PostScriptInstall: $($row.PostScriptInstall)"
#Write-Host "PreScriptUninstall: $($row.PreScriptUninstall)"
#Write-Host "PostScriptUninstall: $($row.PostScriptUninstall)"
#Write-Host "CustomArgumentListInstall: $($row.CustomArgumentListInstall)"
#Write-Host "CustomArgumentListInstall: $($row.CustomArgumentListUninstall)"

#TimeStamp
$Timestamp = (Get-Date).ToString("yyyy-MM-dd-HHmmss")

#Get User
$currentUser = $env:USERNAME

Write-Host "--- Validation Start ---"
write-host "Validating Package $($row.PackageID)"
#Check context is valid
if ($row.Context -notin @("Machine", "machine", "User", "user")) {
    Write-Host "Invalid context setting $($row.Context) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor "Red"
    break
}

#Check AcceptNewerVersion is true or false
if ($row.AcceptNewerVersion -ne $True -and $row.AcceptNewerVersion -ne $False)
{
    Write-Host "Invalid AcceptNewerVersion setting $($row.AcceptNewerVersion) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor "Red"
    break
}

#Check InstallIntent is valid
if ($null -ne $row.InstallIntent -and $row.InstallIntent -ne "") {
    if ($row.InstallIntent -notcontains "Required" -and $row.InstallIntent -notcontains "required" -and $row.InstallIntent -notcontains "Available" -and $row.InstallIntent -notcontains "available")
    {
        Write-Host "Invalid InstallIntent setting $($row.InstallIntent) for package $($row.PackageID) found in CSV. Please review CSV (Use Available or Required) Exiting.." -ForegroundColor "Red"
        break
    }
}

#Check Notification is valid
if ($null -ne $row.Notification -and $row.Notification -ne "") {
    $validNotificationValues = "showAll", "showReboot", "hideAll"
    if ($validNotificationValues -notcontains $row.Notification.ToLower()) {
        Write-Host "Invalid Notification setting $($row.Notification) for package $($row.PackageID) found in CSV. Please review CSV (Use showAll, showReboot, or hideAll) Exiting.." -ForegroundColor "Red"
        break
    }
}

#Check GroupID is set if InstallIntent is specified
if ($null -ne $row.InstallIntent -and $row.InstallIntent -ne "") {
    if ($row.GroupID -eq $null -or $row.GroupID -eq "") {
    Write-Host "Invalid GroupID setting $($row.GroupID) for package $($row.PackageID) found in CSV. Please review CSV. If InstallIntent is set, GroupID must be set too!"
    break
    }
}

#Check UpdateOnly is true or false
Write-Host "Checking UpdateOnly Value for $($row.PackageID)"
if ($row.UpdateOnly -ne $True -and $row.UpdateOnly -ne $False)
{
    Write-Host "Invalid UpdateOnly setting $($row.UpdateOnly) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.."  -ForegroundColor "Red"
    break
}

#Check StopProcessInstall and StopProcessUninstall does not contain "exe"  (inform should not contain .exe)
Write-Host "Checking StopProcessInstall Value for $($row.PackageID)"
if ($row.StopProcessInstall -contains ".") {
    Write-Host "Invalid StopProcessInstall setting $($row.StopProcessInstall) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor "Red"
    break
}
Write-Host "Checking StopProcessUninstall Value for $($row.PackageID)"
if ($row.StopProcessUninstall -contains ".") {
    Write-Host "Invalid StopProcessUninstall setting $($row.StopProcessUninstall) for package $($row.PackageID) found in CSV. Please review the CSV. Exiting.." -ForegroundColor "Red"
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
    Write-Host "Checking Pre-Script and Post-Script Values for $($row.PackageID)"
    $row.PreScriptInstall
    $row.PostScriptInstall
    $row.PreScriptUninstall
    $row.PostScriptUninstall
    if ($row.PreScriptInstall -notlike "*.ps1" -or $row.PostScriptInstall -notlike "*.ps1" -or $row.PreScriptUninstall -notlike "*.ps1" -or $row.PostScriptUninstall -notlike "*.ps1" ) {
        Write-Host "Invalid post or pre-script for package $($row.PackageID) found in CSV. Check that the value contains: .ps1 - Please review the CSV. Exiting.." -ForegroundColor "Red"
        break
}
} 
#>

#Print CustomArgumentListInstall if set and wait confirm
Write-Host "Checking CustomArgumentListInstall Value for $($row.PackageID)"
if ($row.CustomArgumentListInstall -ne "" -or $null)
{
    Write-Host "-- CustomArgumentListInstall --"
    write-host "$($row.CustomArgumentListInstall)"
    if (!$SkipConfirmation) {
    $confirmation = Read-Host "Please confirm CustomArgumentListInstall ($PackageID)? (Y/N)"

    if ($confirmation -eq "Y" -or $confirmation -eq "y") {
     Write-Host "Confirmed"
    } else {
        Write-Host "CustomArgumentListInstall not confirmed. Exiting.."
        return
    }     
}
}

#Print CustomArgumentListUninstall if set and wait confirm
Write-Host "Checking CustomArgumentListUninstall Value for $($row.PackageID)"
if ($row.CustomArgumentListUninstall -ne "" -or $null)
{
    Write-Host "-- CustomArgumentListUninstall --"
    write-host "$($row.CustomArgumentListUninstall)"
    if (!$SkipConfirmation) {
    $confirmation = Read-Host "Please confirm CustomArgumentListUninstall ($PackageID)? (Y/N)"

    if ($confirmation -eq "Y" -or $confirmation -eq "y") {
     Write-Host "Confirmed"
    } else {
        Write-Host "CustomArgumentListUninstall not confirmed. Exiting.."
        return
    }     
}
}

Write-Host "Finished Validation for $($row.PackageID)"
Write-Host "--- Validation End ---"
Write-Host ""

#Print Package details and wait for confirmation. If package not found break.
$PackageIDOutLines = @(winget show --exact --id $($row.PackageID) --scope $($row.Context))
#Check if targetversion specified
if ($null -ne $row.TargetVersion -and $row.TargetVersion -ne "")
{
    $PackageIDOutLines = @(winget show --exact --id $($row.PackageID) --scope $($row.Context) --version $($row.TargetVersion))
}
$PackageIDout = $PackageIDOutLines -join "`r`n"

if ($PackageIDOutLines -notcontains "No package found matching input criteria.") {
    if ($PackageIDOutLines -notcontains "  No applicable installer found; see logs for more details.") {
        if (!$SkipConfirmation) {
        Write-Host "--- WINGET PACKAGE INFORMATION ---"
        Write-Host $PackageIDOut
        Write-Host "--------------------------"
        $confirmation = Read-Host "Confirm the package details above (Y/N)"
        if ($confirmation -eq "N" -or $confirmation -eq "N") {
        break
        }
    }
    } else {
        # Second condition not met
        Write-Host "Applicable installer not found for $($row.Context) context" -ForegroundColor "Red"
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
        $errortext = "Applicable installer not found for $($row.Context) context"
        $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
        continue
    }
} else {
    Write-Host "Package $($row.PackageID) not found using winget" -ForegroundColor "Red"
    return
}


#Scrape Winget package details to use in InTune from $PackageIDOut
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
write-host ""
Write-Host "--- WinGet Scraped Metadata $($row.PackageID) ---"
Write-Host "PackageName: $PackageName"
Write-Host "Version: $Version"
Write-Host "Publisher: $Publisher"
Write-Host "PublisherURL: $PublisherURL"
Write-Host "PublisherSupportURL: $PublisherSupportURL"
Write-Host "Author: $Author"
Write-Host "Description: $Description"
Write-Host "Homepage: $Homepage"
Write-Host "License: $License"
Write-Host "LicenseURL: $LicenseURL"
Write-Host "Copyright: $Copyright"
Write-Host "CopyrightURL: $CopyrightURL"
Write-Host "InstallerType: $InstallerType"
Write-Host "InstallerLocale: $InstallerLocale"
Write-Host "InstallerURL: $InstallerURL"
Write-Host "InstallerSHA256: $InstallerSHA256"
Write-Host ""
Write-Host "--------------------------"

# Build package details to prepare InTune import
$PackageName += " ($($row.Context))"
if ($Row.TargetVersion -ne $null){
    $PackageName += " $($row.TargetVersion)"
}

$PackageName += " (WinGet)"

if ($Row.UpdateOnly -eq $true){
    $PackageName = "Update for " + $PackageName
}

#Clear ArgumentList and CommandLine parameters.
Remove-Variable -Name ArgumentListInstall, ArgumentListUninstall, InstallCommandLine, UninstallCommandLine -ErrorAction SilentlyContinue

#Build ArgumentListInstall (if no custom argumentlist set)
if ([string]::IsNullOrEmpty($row.CustomArgumentListInstall)) {
$ArgumentListInstall = "install --exact --id $($row.PackageID) --silent --accept-package-agreements --accept-source-agreements --scope $($row.Context)"
if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne "") {
$ArgumentListInstall += " --version $($Row.TargetVersion)"
}
}
else {
    $ArgumentListInstall = $row.CustomArgumentListInstall
}
#Replace first 7 characters of ArgumentListInstall with "update" if package is UpdateOnly
if ($Row.UpdateOnly -eq $true) 
{
    $ArgumentListInstall = "update" + $ArgumentListInstall.Substring(7)
}

#Build ArgumentListUninstall (if no custom argumentlist set)
if ([string]::IsNullOrEmpty($row.CustomArgumentListUninstall)) {
$ArgumentListUninstall = "uninstall --exact --id $($row.PackageID) --silent --accept-source-agreements --scope $($row.Context)"
if ($Row.TargetVersion -ne $null -and $Row.TargetVersion -ne "") {
$ArgumentListUninstall += " --version $($Row.TargetVersion)"
}
}
else {
    $ArgumentListUninstall = $row.CustomArgumentListUninstall
}


# Build install commandline
$InstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName $($row.PackageID) -ArgumentList "+[char]34+"$ArgumentListInstall"+[char]34
if ($row.StopProcessInstall -ne $null -and $row.StopProcessInstall -ne ""){
    $InstallCommandLine += " -StopProcess '$($row.StopProcessInstall)'"
}
if ($row.PreScriptInstall -ne $null -and $row.PreScriptInstall -ne ""){
    $InstallCommandLine += " -Prescript '$($row.PreScriptInstall)'"
}

if ($row.PostScriptInstall -ne $null -and $row.PostScriptInstall -ne "") {
    $InstallCommandLine += " -Postscript '$($row.PostScriptInstall)'"
}

# Build uninstall commandline
$UninstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName $($row.PackageID) -ArgumentList "+[char]34+"$ArgumentListUninstall"+[char]34
if ($row.StopProcessUninstall -ne $null -and $row.StopProcessUninstall -ne ""){
    $UninstallCommandLine += " -StopProcess '$($row.StopProcessUninstall)'"
}

if ($row.PreScriptUninstall -ne $null -and $row.PreScriptUninstall -ne "") {
    $InstallCommandLine += " -PreUninstall '$($row.PreScriptUninstall)'"
}

if ($row.PostScriptUninstall -ne $null -and $row.PostScriptUninstall -ne "") {
    $InstallCommandLine += " -PostUninstall '$($row.PostScriptUninstall)'"
}

Write-host " "
write-host "--InstallCommandLine--"
Write-host "$InstallCommandLine"
Write-host " "
write-host "--UninstallCommandLine--"
Write-host "$UninstallCommandLine"
write-host " "

#Make folder to store script files
# Make folder to store script files
$currentDirectory = Get-Location

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



#Build DetectionScript
$WingetDetectionScriptSource = "WinGet-WrapperDetection.ps1"
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
    $settingsSection = $settingsSection -replace '"\$True"','$True'
}
# If $Row.AcceptNewerVersion is $False, replace "$False" with "$False"
elseif ($Row.AcceptNewerVersion -eq $False) {
    $settingsSection = $settingsSection -replace '"\$False"','$False'
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

    Write-Host "Detection Script complete. $WingetDetectionScriptDestination"
} else {
    Write-Host "The # Settings section was not found in the script." -ForegroundColor "Red"
}


#Build RequirementScript
if ($Row.UpdateOnly -eq $True){
    $WingetRequirementScriptSource = "Winget-WrapperRequirements.ps1"
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
        
        Write-Host "Requirement Script complete. $WingetRequirementScriptDestination"
    } else {
        Write-Host "The text 'Exact WinGet Package ID' was not found in the script." -ForegroundColor "Red"
    }
    }

#Build Targetversion to use with intune package
if ($row.TargetVersion -ne $null -and $row.TargetVersion -ne "") {
    $VersionInfo = $row.TargetVersion
    }
else {
    $VersionInfo = "Latest"
}

#Convert Context
if ($row.Context -in @("Machine", "machine")) {
    $ContextConverted = "System"
}
elseif ($row.Context -in @("User", "user")) {
    $ContextConverted = "User"
}

#Build IntuneWin Folder  (including pre/postscripts if specified in CSV)
# Create a folder named "InTuneWin" if it doesn't exist
try {
$intuneWinFolderPath = Join-Path -Path $PackageFolderPath -ChildPath "InTuneWin"
if (-not (Test-Path -Path $intuneWinFolderPath -PathType Container)) {
    New-Item -Path $intuneWinFolderPath -ItemType Directory | Out-Null
}
# Copy WinGet-Wrapper.ps1
Copy-Item -Path "WinGet-Wrapper.ps1" -Destination $intuneWinFolderPath -ErrorAction Stop

# Copy pre or post script if specified
if ($row.PreScriptInstall -ne $null -and $row.PreScriptInstall -ne "") {
    Copy-Item -Path $row.PreScriptInstall -Destination $intuneWinFolderPath -ErrorAction Stop
    Write-Host "Copied $($row.PreScriptInstall) to $($intuneWinFolderPath)"
}
if ($row.PostScriptInstall -ne $null -and $row.PostScriptInstall -ne "") {
    Copy-Item -Path $row.PostScriptInstall -Destination $intuneWinFolderPath -ErrorAction Stop
    Write-Host "Copied $($row.PostScriptInstall) to $($intuneWinFolderPath)"
}
if ($row.PreScriptUninstall -ne $null -and $row.PreScriptUninstall -ne "") {
    Copy-Item -Path $row.PreScriptUninstall -Destination $intuneWinFolderPath -ErrorAction Stop
    Write-Host "Copied $($row.PreScriptUninstall) to $($intuneWinFolderPath)"
}
if ($row.PostScriptUninstall -ne $null -and $row.PostScriptUninstall -ne "") {
    Copy-Item -Path $row.PostScriptUninstall -Destination $intuneWinFolderPath -ErrorAction Stop
    Write-Host "Copied $($row.PostScriptUninstall) to $($intuneWinFolderPath)"
}}
catch {
    Write-host "Failed to copy files to build InTuneWin package. Please check pre/post script is found"
    return
}

# Build WinGet-Wrapper.intunewin
Write-Host "Building WinGet-Wrapper.InTuneWin file"
try {
[string]$toolargs = "-c ""$($intuneWinFolderPath)"" -s ""WinGet-Wrapper.ps1"" -o ""$($PackageFolderPath)"" -q"
(Start-Process -FilePath "$PSScriptRoot\IntuneWinAppUtil.exe" -ArgumentList $toolargs -PassThru:$true -ErrorAction Stop -NoNewWindow).WaitForExit()
    
    # Check that IntuneWin was created
    if (Test-Path -Path "$PackageFolderPath\WinGet-Wrapper.intunewin" -PathType Leaf) {
        Write-host "IntuneWinAppUtil.exe built WinGet-Wrapper.intunewin for $PackageName"
    }
    else {
        Write-host "IntuneWinAppUtil.exe could not build WinGet-Wrapper.intunewin"
        return
    }
    
}
catch {
    write-host "Error running IntuneWinAppUtil.exe"
    return
}

#Set WinGet-Wrapper.intunewin path
$WinGetWrapperInTuneWinFilePath = Join-Path -Path $PackageFolderPath -ChildPath 'WinGet-Wrapper.intunewin'

#Detection
$DetectionRuleScript = New-IntuneWin32AppDetectionRuleScript -ScriptFile $WingetDetectionScriptDestination

#RequirementRule Base
$RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "All" -MinimumSupportedWindowsRelease "W10_20H2" 

# Build base Add-InTuneWin32App Arguments
$AddInTuneWin32AppArguments = @{
    FilePath = "$WinGetWrapperInTuneWinFilePath"
    DisplayName = "$PackageName"
    Description = "$Description"
    Publisher = "$Publisher"
    AppVersion = "$VersionInfo"
    Developer = "$Author"
    InstallCommandLine = "$InstallCommandLine"
    UninstallCommandLine = "$UninstallCommandLine"
    InstallExperience = "$ContextConverted"
    RestartBehavior = "suppress"
    Owner = "$currentUser"
    Notes = "Created by $currentUser at $Timestamp using WinGet-WrapperCreateFromCSV (https://github.com/SorenLundt/WinGet-Wrapper)"
    DetectionRule = $DetectionRuleScript
    RequirementRule = $RequirementRule
}

# Build AdditionalRequirementRule (if UpdateOnly True)
if ($Row.UpdateOnly -eq $True) {
    $AdditionalRequirementRule = New-IntuneWin32AppRequirementRuleScript -ScriptFile $WingetRequirementScriptDestination -ScriptContext "$ContextConverted" -StringComparisonOperator "equal" -StringOutputDataType -StringValue "Installed"
    $AddInTuneWin32AppArguments.AdditionalRequirementRule = $AdditionalRequirementRule
} else {
}

# If overwrite enabled , remove win32app from intune before importing
# Disabled section. Needs more work.
<#
if ($Overwrite -eq $True) {
    try{
    write-host "Overwrite enabled - Attempting to remove Win32App if it already exists in InTune - $PackageName "
    write-host "$packagename"
    $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName"
    write-host "627 - $CheckIntuneAppExists"
    if ($CheckIntuneAppExists.Count -gt 0)
    {
    write-host "Package found in InTune "
    Remove-IntuneWin32App -DisplayName "$PackageName"
    [int]$NumberOfTimesChecked = 1
    do {
        write-host "Checking if Win32App was removed before continuing. Checks: $NumberOfTimesChecked"
        write-host "635 - $CheckIntuneAppExists"
        $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName"
        Start-Sleep 15
        $NumberOfTimesChecked++
    } while ($CheckIntuneAppExists.Count -gt 0)
    }
}
    catch {
        Write-Host "An error occurred: $_.Exception.Message" -ForeGroundColor "Red"
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
        $errortext = "Error Remvoving Win32App: $_.Exception.Message"
        $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
        continue
    }

}
#>
    

Write-Host "Checking if application already exists in Intune - $PackageName"
    # Get the Intune Win32 apps with the specified display name
    $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName" -WarningAction SilentlyContinue

    # Check if any matching applications were found
    write-host "CheckInTuneAppExists.Count = $($CheckIntuneAppExists.Count)"
    if ($CheckIntuneAppExists.Count -gt 0) {
        Write-Host "ERROR: Application with the same name already exists in Intune. Could not import - $PackageName" -ForegroundColor "Red"
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
        $errortext = "Already exists in InTune"
        $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
        continue
    }
    else {
        Write-Host "OK! A similar package was not found in Intune."
    }

#Import application to InTune
try {
    Write-Host "Importing application '$PackageName' to InTune"
    $AppImport = Add-IntuneWin32App @AddInTuneWin32AppArguments -WarningAction Continue -ErrorAction Continue
    $AppID = $AppImport.ID
    $row | Add-Member -MemberType NoteProperty -Name "AppID" -Value $AppID #Write AppID to $row
    Write-Host "Imported application '$PackageName' to InTune" -ForeGroundColor "Green"
    $imported = $True
}
catch {
    Write-Host "An error occurred: $_.Exception.Message" -ForeGroundColor "Red"
    $imported = $False
    $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
    $errortext = "Error Importing: $_.Exception.Message"
    $row | Add-Member -MemberType NoteProperty -Name "ErrorText" -Value $errortext  #Write errortext to $row
    continue
}

# Deploy Application if set - Install Intent "Required"
try {
    if ($Row.InstallIntent -contains "Required" -and ($null -ne $Row.GroupID -and $Row.GroupID -ne "")) {
        Add-IntuneWin32AppAssignmentGroup -Include -ID $Row.AppID -GroupID $Row.GroupID -Intent $row.InstallIntent -Notification $row.Notification -ErrorAction continue
        Write-Host "Deployed AppID:$($Row.AppID) to $($Row.GroupID)" -ForeGroundColor "Green"
    }
    
    # Deploy Application if set - Install Intent "Available"
    if ($Row.InstallIntent -contains "Available" -and ($null -ne $Row.GroupID -and $Row.GroupID -ne "")) {
        Add-IntuneWin32AppAssignmentGroup -Include -ID $Row.AppID -GroupID $Row.GroupID -Intent $row.InstallIntent -Notification $row.Notification -ErrorAction continue
        Write-Host "Deployed AppID:$($Row.AppID) to $($Row.GroupID)" -ForeGroundColor "Green"
    }
}
 catch {
    Write-Host "Error deploying $($Row.PackageID) : $_"
}

#Success
$imported = $True
$row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row

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
        write-host "Waiting 15s before importing next package.. (Throttle)"
        Start-Sleep -Seconds 15
    }

}
}
###### END FOREACH ######

#Write Results
Write-host "---- RESULTS Package Creation ----"
foreach ($row in $data) {
    $importedStatus = $row.Imported
    $textColor = "Green"  # Default to green
    $importedtext = "Success" #Default to success
    if (-not $importedStatus) {
        $textColor = "Red"  # Change to red if Imported is False
        $importedtext = "Failed"
    }
    write-host ""
    $formattedText = "IMPORTED:$ImportedText PackageID:$($row.PackageID) AppID:$($row.AppID) TargetVersion:$($row.TargetVersion) Context:$($row.Context) UpdateOnly: $($row.UpdateOnly) AcceptNewerVersion: $($row.AcceptNewerVersion) ErrorText: $($row.ErrorText)"
    Write-Host $formattedText -ForegroundColor $textColor

    # If deployed also show these results
    if ($null = $row.GroupID -or $row.GroupID -ne ""){
        if ($importedStatus -eq $True)
        {
        Write-host "> DEPLOYED:$($row.PackageID) AppID:$($row.AppID) ----> GroupID:$($row.GroupID) InstallIntent:$($row.InstallIntent) Notification:$($row.Notification)" -ForegroundColor $textColor
        }   
    }

}