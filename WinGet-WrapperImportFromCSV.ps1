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
# PackageID,Context,AcceptNewerVersion,UpdateOnly,TargetVersion,StopProcessInstall,StopProcessUninstall,PreScriptInstall,PostScriptInstall,PreScriptUninstall,PostScriptUninstall,CustomArgumentListInstall,CustomArgumentListUninstall
#
# Requirements:
# Requires Script files and IntuneWinAppUtil.exe to be present in script directory
#
# Version History
# Version 1.0 - 12-09-2023 SorenLundt - Initial Version

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
    Write-Host "File does not found: $csvFile" -ForegroundColor "Red"
    return
}

# Import the CSV file with custom headers
$data = Import-Csv -Path $csvFile -Header "PackageID", "Context", "AcceptNewerVersion", "UpdateOnly", "TargetVersion", "StopProcessInstall", "StopProcessUninstall", "PreScriptInstall", "PostScriptInstall", "PreScriptUninstall", "PostScriptUninstall", "CustomArgumentListInstall", "CustomArgumentListUninstall" | Select-Object -Skip 1

# Convert "AcceptNewerVersion" and "UpdateOnly" columns to Boolean values
$data = $data | ForEach-Object {
    $_.AcceptNewerVersion = [bool]($_.AcceptNewerVersion -as [int])
    $_.UpdateOnly = [bool]($_.UpdateOnly -as [int])
    $_
}
Write-host "-- Packagelist to Import --"
foreach ($row in $data){
    Write-Host "$($row.PackageID)"
}
write-host ""

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
        Write-Host "--- PACKAGE INFORMATION ---"
        Write-Host $PackageIDOut
        Write-Host "--------------------------"
        if (!$SkipConfirmation) {
        $confirmation = Read-Host "Confirm the package details above (Y/N)"
        if ($confirmation -eq "N" -or $confirmation -eq "N") {
        break
        }
        }
    } else {
        # Second condition not met
        Write-Host "Applicable installer not found for $($row.Context) context" -ForegroundColor "Red"
        return
    }
} else {
    Write-Host "Package $($row.PackageID) not found using winget" -ForegroundColor "Red"
    return
}

write-host ""
Write-Host "--- Scrape WinGet Details $($row.PackageID) ---"
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
$InstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName '$($row.PackageID)' -ArgumentList '$ArgumentListInstall'"
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
$UninstallCommandLine = "Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File WinGet-Wrapper.ps1 -PackageName '$($row.PackageID)' -ArgumentList '$ArgumentListUninstall'"
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
    New-Item -Path $packagesDirectory -ItemType Directory
}

# Combine the "Packages" directory path with the folder name to create the full path
$PackageFolderPath = Join-Path -Path $packagesDirectory -ChildPath $folderName

# Create the subfolder with the desired name
New-Item -Path $PackageFolderPath -ItemType Directory



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
# If $Row.AcceptNewerVersion is $False, replace "$True" with "$False"
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
    New-Item -Path $intuneWinFolderPath -ItemType Directory
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
try {
[string]$toolargs = "-c ""$($intuneWinFolderPath)"" -s ""WinGet-Wrapper.ps1"" -o ""$($PackageFolderPath)"" -q"
write-host $toolargs
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

Write-Host "Checking if Application '$PackageName' already exists in Intune"
    # Get the Intune Win32 apps with the specified display name
    $CheckIntuneAppExists = Get-IntuneWin32App -DisplayName "$PackageName" -WarningAction SilentlyContinue

    # Check if any matching applications were found
    if ($CheckIntuneAppExists.Count -gt 0) {
        Write-Host "ERROR: Application with the same name ($PackageName) already exists in Intune. Could not import $($Row.PackageID)" -ForegroundColor "Red"
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
        continue
    }
    else {
        Write-Host "OK! A similar package was not found in Intune."
    }

#Import application to InTune
try {
    Write-Host "Importing application '$PackageName' to InTune"
    Add-IntuneWin32App @AddInTuneWin32AppArguments
    Write-Host "Imported application '$PackageName' to InTune" -ForeGroundColor "Green"
    $imported = $True
}
catch {
    Write-Host "An error occurred: $_.Exception.Message" -ForeGroundColor "Red"
    $imported = $False
    $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row
    continue
}

#Success
$imported = $True
$row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row

    }
    catch {
        #Failed
        $imported = $False
        $row | Add-Member -MemberType NoteProperty -Name "Imported" -Value $imported  #Write imported status to $row

    }
}

#If more packages to import, wait 30 seconds to avoid creating issue (InTune seems to maybe have a throttle in place..)
    # Check if there are more than 2 rows in $data
    if ($data.Count -gt 2) {
        write-host "Waiting 30s before importing next package..InTune Throttling"
        Start-Sleep -Seconds 30
    }
}
###### END FOREACH ######

#Write Results
Write-host "---- RESULTS ----"
foreach ($row in $data) {
    $importedStatus = $row.Imported
    $textColor = "Green"  # Default to green
    $importedtext = "Success" #Default to success
    if (-not $importedStatus) {
        $textColor = "Red"  # Change to red if Imported is False
        $importedtext = "Failed"
    }
    $formattedText = "Imported:$ImportedText PackageID:$($row.PackageID) TargetVersion: $($row.TargetVersion) Context:$($row.Context) UpdateOnly: $($row.UpdateOnly) AcceptNewerVersion: $($row.AcceptNewerVersion)"
    Write-Host $formattedText -ForegroundColor $textColor
}