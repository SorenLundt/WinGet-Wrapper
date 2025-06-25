# Soren Lundt - 12-02-2024
# URL: https://github.com/SorenLundt/WinGet-Wrapper
# License: https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/LICENSE.txt
# Graphical interface for WinGet-Wrapper Import
# Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm

# Requirements:
# Requires Script files and IntuneWinAppUtil.exe to be present in script directory
#
# Version History
# Version 1.0 - 12-02-2024 SorenLundt - Initial Version
# Version 1.1 - 21-02-2024 SorenLundt - Fixed issue where only 1 package was imported to InTune (Script assumed there was just one row)
# Version 1.2 - 30-10-2024 SorenLundt - Various improvements (Ability to get details and available versions for packages, loading GUI progressbar, link to repo, unblock script files, bug fixes.)
# Version 1.3 - 31-10-2024 SorenLundt - Added Column Definitions help button, added SKIPMODULECHECK param to skip check if required modules are up-to-date. (testing use)

#Parameters
Param (
    #Skip module checking (for testing purposes)
    [Switch]$SKIPMODULECHECK = $false
)

# Greeting
Write-Host ''
Write-Host '****************************************************'
Write-Host '                  WinGet-Wrapper)'
Write-Host '  https://github.com/SorenLundt/WinGet-Wrapper'
Write-Host ''
Write-Host '          GNU General Public License v3'
Write-Host '****************************************************'
Write-Host '   WinGet-WrapperImportGUI Starting up..'
Write-Host ''

# Function to show loading progress bar in console.
function Show-ConsoleProgress {
    param (
        [string]$Activity = 'Loading Winget-Wrapper Import GUI',
        [string]$Status = '',
        [int]$PercentComplete = 0
    )
    Write-Host "$Status - [$PercentComplete%]"
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 0 -Status 'Initializing...'

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set the timestamp for log file
$timestamp = Get-Date -Format 'yyyyMMddHHmmss'

#Find Script root path  
if (-not $PSScriptRoot) {
    $scriptRoot = (Get-Location -PSProvider FileSystem).ProviderPath
}
else {
    $scriptRoot = $PSScriptRoot
}

# Create logs folder if it doesn't exist
$LogFolder = Join-Path -Path $scriptRoot -ChildPath 'Logs'
# Create logs folder if it doesn't exist
if (-not (Test-Path -Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

# Install and load required modules
$intuneWin32AppModule = 'IntuneWin32App'
$microsoftGraphIntuneModule = 'Microsoft.Graph.Intune'
$microsoftGraphAuthenticationModule = 'Microsoft.Graph.Authentication'

#DEBUG (Skip ModuleCheck)
#$SKIPMODULECHECK = $true 
if (-not $SKIPMODULECHECK) {

    # Check IntuneWin32App module
    # Update ConsoleProgress
    Show-ConsoleProgress -PercentComplete 10 -Status "Checking and updating $intuneWin32AppModule.."  
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
            Write-Host "Module $intuneWin32AppModule is already up-to-date." -ForegroundColor Green
        }
    }
    # Check Microsoft.Graph.Intune module
    # Update ConsoleProgress
    Show-ConsoleProgress -PercentComplete 40 -Status "Checking and updating $microsoftGraphIntuneModule.."  
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
            Write-Host "Module $microsoftGraphIntuneModule is already up-to-date." -ForegroundColor Green
        }
    }


    # Check Microsoft.Graph.Authentication module
    # Update ConsoleProgress
    Show-ConsoleProgress -PercentComplete 40 -Status "Checking and updating $microsoftGraphAuthenticationModule.."  
    $moduleInstalled = Get-InstalledModule -Name $microsoftGraphAuthenticationModule -ErrorAction SilentlyContinue

    if (-not $moduleInstalled) {
        Install-Module -Name $microsoftGraphAuthenticationModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $microsoftGraphAuthenticationModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $microsoftGraphAuthenticationModule -Force
        }
        else {
            Write-Host "Module $microsoftGraphAuthenticationModule is already up-to-date." -ForegroundColor Green

        }
    }
}


#Import modules
Show-ConsoleProgress -PercentComplete 60 -Status "Importing module $intuneWin32AppModule.."
Import-Module -Name 'IntuneWin32App'

Show-ConsoleProgress -PercentComplete 80 -Status "Importing module $microsoftGraphIntuneModule.."
Import-Module -Name 'Microsoft.Graph.Intune'

Show-ConsoleProgress -PercentComplete 90 -Status 'Unblocking script files (Unblock-File)'
# Unblock all files in the current directory
$files = Get-ChildItem -Path . -File
foreach ($file in $files) {
    try {
        Unblock-File -Path $file.FullName
    }
    catch {
        Write-Host "Failed to unblock: $($file.FullName) - $_"
    }
}

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 80 -Status 'Loading functions..'   

#Functions
function Write-ConsoleTextBox {
    param (
        [string]$Message,
        [switch]$NoTimeStamp
    )

    if (-not $NoTimeStamp) {
        $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $Message = "[$TimeStamp] $Message"
    }

    # Output to the console
    Write-Host $Message
    
    # Append to the console textbox
    $consoleTextBox.AppendText("$Message`r`n")
}


# Function to read log file and update the GUI
function Update-GUIFromLogFile {
    param (
        [string]$logFilePath
    )
    # Read the log file content
    $logContent = Get-Content -Path $logFilePath

    # Assuming Write-ConsoleTextBox adds each line to the GUI's textbox
    foreach ($line in $logContent) {
        Write-ConsoleTextBox -Message $line -NoTimeStamp
    }
}

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 90 -Status 'Loading GUI elements..'   

# GUI
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Winget-Wrapper Import GUI  - https://github.com/SorenLundt/WinGet-Wrapper'
$form.Width = 1475
$form.Height = 980
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.Topmost = $false
$form.MaximizeBox = $False

# Set the icon for the form
Write-Host "$scriptRoot"
$iconPath = Join-Path -Path $scriptRoot -ChildPath 'Winget-Wrapper.ico'
if (Test-Path $iconPath) {
    $icon = New-Object System.Drawing.Icon($iconPath)
    $form.Icon = $icon
}
else {
    Write-Host "Icon file not found at $iconPath"
}

<# Create a Link to GitHub page
$linkLabelGitHub = New-Object System.Windows.Forms.LinkLabel
$linkLabelGitHub.Text = 'Visit the GitHub Repository'
$linkLabelGitHub.Location = New-Object System.Drawing.Point(1296, 5)  # Adjusted position for better layout
$linkLabelGitHub.AutoSize = $true
$linkLabelGitHub.LinkColor = [System.Drawing.Color]::Blue
$linkLabelGitHub.Font = New-Object System.Drawing.Font('Arial', 8, [System.Drawing.FontStyle]::Bold)  # Use the same font as the Column Guide
$linkLabelGitHub.BackColor = [System.Drawing.Color]::LightGray  # Background color for consistency
$linkLabelGitHub.Padding = New-Object System.Windows.Forms.Padding(1)  # Add padding for consistency
$linkLabelGitHub.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle  # Add a border to make it stand out
$form.Controls.Add($linkLabelGitHub)#>

# Create a Link to GitHub page
$linkLabelGitHub = New-Object System.Windows.Forms.LinkLabel
$linkLabelGitHub.Text = 'Visit the GitHub Repository'
$linkLabelGitHub.Location = New-Object System.Drawing.Point(1280, 10)
$linkLabelGitHub.AutoSize = $true
$linkLabelGitHub.LinkColor = [System.Drawing.Color]::Blue
$linkLabelGitHub.Font = New-Object System.Drawing.Font('Arial', 10, [System.Drawing.FontStyle]::Underline)
$form.Controls.Add($linkLabelGitHub)

# Create a help label to get column definitions
$ColumnHelpLabel = New-Object System.Windows.Forms.Label
$ColumnHelpLabel.Text = 'Show Column Definitions'
$ColumnHelpLabel.Location = New-Object System.Drawing.Point(1310, 40)
$ColumnHelpLabel.AutoSize = $true
#$ColumnHelpLabel.ForeColor = [System.Drawing.Color]::Blue  # Set text color to blue
$ColumnHelpLabel.Font = New-Object System.Drawing.Font('Arial', 8, [System.Drawing.FontStyle]::Bold)  # Make the text bold
$ColumnHelpLabel.BackColor = [System.Drawing.Color]::LightGray  # Background color for contrast
$ColumnHelpLabel.Padding = New-Object System.Windows.Forms.Padding(1)  # Add some padding around the text
$form.Controls.Add($ColumnHelpLabel)

# Add click event to the LinkLabel
$linkLabelGitHub.Add_LinkClicked({
        Start-Process 'https://github.com/SorenLundt/WinGet-Wrapper'  # Change to your desired URL
    })

# Click event for Column Help Button
$ColumnHelpLabel.Add_Click({
        GetColumnDefinitions
    })

# Create TextBox for search string
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10, 10)
$searchBox.Width = 400
$form.Controls.Add($searchBox)
$searchBoxtooltip = New-Object System.Windows.Forms.ToolTip
$searchBoxtooltip.SetToolTip($searchBox, 'Enter software name, ex. VLC, 7-zip, etc.')

# Create Button for search
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(420, 8)
$searchButton.Width = 50
$searchButton.Text = 'Search'
$form.Controls.Add($searchButton)

# Create Label to display search error
$searchErrorLabel = New-Object System.Windows.Forms.Label
$searchErrorLabel.Location = New-Object System.Drawing.Point(480, 10)
$searchErrorLabel.Width = 500
#$searchErrorLabel.ForeColor = [System.Drawing.Color]::Tomato
$form.Controls.Add($searchErrorLabel)

# Create Label for $dataGridView (Results)
$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Location = New-Object System.Drawing.Point(10, 37)
$resultsLabel.Width = 200
$resultsLabel.Text = 'WinGet Packages'
$form.Controls.Add($resultsLabel)

# Create DataGridView for results
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 60)
$dataGridView.Width = 600
$dataGridView.Height = 500
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect  # Only allow full row selection
$form.Controls.Add($dataGridView)

# Add columns to DataGridView
$dataGridView.Columns.Add('Name', 'Name')
$dataGridView.Columns.Add('ID', 'ID')
$dataGridView.Columns.Add('Version', 'Version')

# Set initial widths for columns in the DataGridView
$dataGridView.Columns['Name'].Width = 200  # Adjust the width as needed
$dataGridView.Columns['ID'].Width = 150    # Adjust the width as needed
$dataGridView.Columns['Version'].Width = 100  # Adjust the width as needed

# Create Label for $dataGridView (Selected)
$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Location = New-Object System.Drawing.Point(650, 37)
$resultsLabel.Width = 200
$resultsLabel.Text = 'InTune Import List'
$form.Controls.Add($resultsLabel)

# Create a second DataGridView
$dataGridViewSelected = New-Object System.Windows.Forms.DataGridView
$dataGridViewSelected.Location = New-Object System.Drawing.Point(650, 60)
$dataGridViewSelected.Width = 800
$dataGridViewSelected.Height = 500
$form.Controls.Add($dataGridViewSelected)

# Add columns to the second DataGridView
$dataGridViewSelected.Columns.Add('PackageID', 'PackageID')
$dataGridViewSelected.Columns.Add('Context', 'Context')
$dataGridViewSelected.Columns.Add('AcceptNewerVersion', 'AcceptNewerVersion')
$dataGridViewSelected.Columns.Add('UpdateOnly', 'UpdateOnly')
$dataGridViewSelected.Columns.Add('TargetVersion', 'TargetVersion')
$dataGridViewSelected.Columns.Add('StopProcessInstall', 'StopProcessInstall')
$dataGridViewSelected.Columns.Add('StopProcessUninstall', 'StopProcessUninstall')
$dataGridViewSelected.Columns.Add('PreScriptInstall', 'PreScriptInstall')
$dataGridViewSelected.Columns.Add('PostScriptInstall', 'PostScriptInstall')
$dataGridViewSelected.Columns.Add('PreScriptUninstall', 'PreScriptUninstall')
$dataGridViewSelected.Columns.Add('PostScriptUninstall', 'PostScriptUninstall')
$dataGridViewSelected.Columns.Add('CustomArgumentListInstall', 'CustomArgumentListInstall')
$dataGridViewSelected.Columns.Add('CustomArgumentListUninstall', 'CustomArgumentListUninstall')
$dataGridViewSelected.Columns.Add('InstallIntent', 'InstallIntent')
$dataGridViewSelected.Columns.Add('Notification', 'Notification')
$dataGridViewSelected.Columns.Add('GroupID', 'GroupID')

# Set initial widths for columns in the DataGridViewSelected
foreach ($column in $dataGridViewSelected.Columns) {
    $column.Width = '80'
}

# Create Button for exporting CSV from dataGridViewSelected
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(760, 565)
$exportButton.Width = 100
$exportButton.Text = 'Export CSV'
$form.Controls.Add($exportButton)

# Create Button for moving selected rows with a right-arrow icon
$moveButton = New-Object System.Windows.Forms.Button
$moveButton.Location = New-Object System.Drawing.Point(618, 280)
$moveButton.Width = 30
$moveButton.Height = 30
$moveButtontooltip = New-Object System.Windows.Forms.ToolTip
$moveButtontooltip.SetToolTip($moveButton, 'Add Selected Package(s) to Import List')

# Create a Bitmap for the arrow icon
$arrowIcon = New-Object System.Drawing.Bitmap 50, 50

# Draw a right arrow onto the Bitmap
$arrowGraphics = [System.Drawing.Graphics]::FromImage($arrowIcon)
$arrowBrush = [System.Drawing.Brushes]::Black
$arrowGraphics.FillPolygon($arrowBrush, @(
        [System.Drawing.Point]::new(10, 10),
        [System.Drawing.Point]::new(30, 25),
        [System.Drawing.Point]::new(10, 40)
    ))
$arrowGraphics.Dispose()

# Set the button's appearance and icon
$moveButton.FlatStyle = 'Flat'
$moveButton.FlatAppearance.BorderSize = 0
$moveButton.BackgroundImage = $arrowIcon
$moveButton.BackgroundImageLayout = 'Stretch'
$moveButton.Text = ''
$form.Controls.Add($moveButton)

# Create Button for deleting selected rows
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Location = New-Object System.Drawing.Point(1350, 565)
$deleteButton.Width = 100
$deleteButton.Text = 'Delete Selected'
$form.Controls.Add($deleteButton)

# Create Button for importing CSV to dataGridViewSelected
$importCSVButton = New-Object System.Windows.Forms.Button
$importCSVButton.Location = New-Object System.Drawing.Point(650, 565)
$importCSVButton.Width = 100
$importCSVButton.Text = 'Import CSV'
$form.Controls.Add($importCSVButton)

# Create a TextBox for console output
$consoleTextBox = New-Object System.Windows.Forms.TextBox
$consoleTextBox.Location = New-Object System.Drawing.Point(10, 680)
$consoleTextBox.Width = 1420
$consoleTextBox.Height = 280
$consoleTextBox.Multiline = $true
$consoleTextBox.ScrollBars = 'Vertical'
$form.Controls.Add($consoleTextBox)

# Create Button for importing data into InTune
$InTuneimportButton = New-Object System.Windows.Forms.Button
$InTuneimportButton.Location = New-Object System.Drawing.Point(650, 640)
$InTuneimportButton.Width = 120
$InTuneimportButton.Text = 'Import to InTune'
$IntuneImportButton.Visible = $True
$form.Controls.Add($InTuneimportButton)

# Create Label for $tenantId
$tenantIdLabel = New-Object System.Windows.Forms.Label
$tenantIdLabel.Location = New-Object System.Drawing.Point(650, 600)
$tenantIdLabel.Width = 200
$tenantIdLabel.Height = 15
$tenantIdLabel.Text = 'Tenant ID'
$form.Controls.Add($tenantIdLabel)

# Create TextBox for Tenant ID
$tenantIDTextBox = New-Object System.Windows.Forms.TextBox
$tenantIDTextBox.Location = New-Object System.Drawing.Point(650, 615)
$tenantIDTextBox.Width = 250
$form.Controls.Add($tenantIDTextBox)

# Help text for Tenant ID textbox
$tenantIDTextBoxDefaultText = 'Enter Tenant ID (e.g., company.onmicrosoft.com)'
$tenantIDTextBox.Text = "$tenantIDTextBoxDefaultText"
$tenantIDTextBox.ForeColor = [System.Drawing.Color]::Gray

# Create Label for $clientId
$clientIdLabel = New-Object System.Windows.Forms.Label
$clientIdLabel.Location = New-Object System.Drawing.Point(920, 600)
$clientIdLabel.Width = 50
$clientIdLabel.Height = 15
$clientIdLabel.Text = 'Client ID'
$form.Controls.Add($clientIdLabel)

$clientIdDefaultLabel = New-Object System.Windows.Forms.Label
$clientIdDefaultLabel.Location = New-Object System.Drawing.Point(970, 600)
$clientIdDefaultLabel.Width = 230
$clientIdDefaultLabel.Height = 15
$clientIdDefaultLabel.Font = New-Object System.Drawing.Font('Arial', 8)
$clientIdDefaultLabel.ForeColor = [System.Drawing.Color]::Gray
$clientIdDefaultLabel.Text = 'Default: Microsoft Graph Command Line Tools'
$form.Controls.Add($clientIdDefaultLabel)


# Create TextBox for Client ID
$clientIDTextBox = New-Object System.Windows.Forms.TextBox
$clientIDTextBox.Location = New-Object System.Drawing.Point(920, 615)
$clientIDTextBox.Width = 280
$form.Controls.Add($clientIDTextBox)

# Help text for Client textbox
$clientIDTextBox.Text = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
#$clientIDTextBox.ForeColor = [System.Drawing.Color]::Gray

# Create Label for $redirectURI
$redirectLabel = New-Object System.Windows.Forms.Label
$redirectLabel.Location = New-Object System.Drawing.Point(1210, 600)
$redirectLabel.Width = 190
$redirectLabel.Height = 15
$redirectLabel.Text = 'Redirect URI'
$form.Controls.Add($redirectLabel)

# Create TextBox for Redirect URL
$redirectURLTextBox = New-Object System.Windows.Forms.TextBox
$redirectURLTextBox.Location = New-Object System.Drawing.Point(1210, 615)
$redirectURLTextBox.Width = 200
$form.Controls.Add($redirectURLTextBox)

# Help text for Redirect url textbox
$redirectURLTextBox.Text = 'https://login.microsoftonline.com/common/oauth2/nativeclient'
#$redirectURLTextBox.ForeColor = [System.Drawing.Color]::Gray

# Create Button for Getting Package Details
$GetPackageDetails = New-Object System.Windows.Forms.Button
$GetPackageDetails.Location = New-Object System.Drawing.Point(10, 565)
$GetPackageDetails.Width = 240
$GetPackageDetails.Text = 'Get detailed info for selected package(s)'
$GetPackageDetails.Visible = $True
$form.Controls.Add($GetPackageDetails)

# Create Button for Getting Package Versions
$GetPackageVersions = New-Object System.Windows.Forms.Button
$GetPackageVersions.Location = New-Object System.Drawing.Point(260, 565)
$GetPackageVersions.Width = 240
$GetPackageVersions.Text = 'Get available versions for selected package(s)'
$GetPackageVersions.Visible = $True
$form.Controls.Add($GetPackageVersions)

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 95 -Status 'Loading Event handlers..'

# EVENTS
# Event handler for when the textbox gains focus (Enter event)
$tenantIDTextBox.Add_Enter({
        if ($tenantIDTextBox.Text -eq "$tenantIDTextBoxDefaultText") {
            $tenantIDTextBox.Text = ''
            $tenantIDTextBox.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Event handler for when the textbox loses focus (Leave event)
$tenantIDTextBox.Add_Leave({
        if ([string]::IsNullOrWhiteSpace($tenantIDTextBox.Text)) {
            $tenantIDTextBox.Text = "$tenantIDTextBoxDefaultText"
            $tenantIDTextBox.ForeColor = [System.Drawing.Color]::Gray
        }
    })

# Event handler for when the textbox gains focus (Enter event)
$clientIDTextBox.Add_Enter({
        if ($clientIDTextBox.Text -eq "$clientIDTextBoxDefaultText") {
            $clientIDTextBox.Text = ''
            $clientIDTextBox.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Event handler for when the textbox loses focus (Leave event)
$clientIDTextBox.Add_Leave({
        if ([string]::IsNullOrWhiteSpace($clientIDTextBox.Text)) {
            $clientIDTextBox.Text = "$clientIDTextBoxDefaultText"
            $clientIDTextBox.ForeColor = [System.Drawing.Color]::Gray
        }
    })

# Event handler for when the textbox gains focus (Enter event)
$redirectURLTextBox.Add_Enter({
        if ($redirectURLTextBox.Text -eq "$redirectURLTextBoxDefaultText") {
            $redirectURLTextBox.Text = ''
            $redirectURLTextBox.ForeColor = [System.Drawing.Color]::Black
        }
    })

# Event handler for when the textbox loses focus (Leave event)
$redirectURLTextBox.Add_Leave({
        if ([string]::IsNullOrWhiteSpace($redirectURLTextBox.Text)) {
            $redirectURLTextBox.Text = "$redirectURLTextBoxDefaultText"
            $redirectURLTextBox.ForeColor = [System.Drawing.Color]::Gray
        }
    })

# Event handler for deleting selected rows
$deleteButton.Add_Click({
        foreach ($row in $dataGridViewSelected.SelectedRows) {
            $packageID = $row.Cells['PackageID'].Value
            $dataGridViewSelected.Rows.Remove($row)
            Write-ConsoleTextBox "Removed '$PackageID' from import list"

        }

        # Autosize columns after deleting rows
        $dataGridViewSelected.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
    })
# Event handler for importing CSV
$importCSVButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = (Get-Location).Path
        $openFileDialog.Filter = 'CSV files (*.csv)|*.csv'
    
        $result = $openFileDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $csvFilePath = $openFileDialog.FileName
        
            # Load CSV data into a PowerShell object
            $importedData = Import-Csv -Path $csvFilePath
        
            # Clear existing rows in $dataGridViewSelected
            $dataGridViewSelected.Rows.Clear()
        
            # Add rows to $dataGridViewSelected from imported data
            foreach ($row in $importedData) {
                $dataGridViewSelected.Rows.Add(
                    $row.PackageID, 
                    $row.Context, 
                    $row.AcceptNewerVersion, 
                    $row.UpdateOnly, 
                    $row.TargetVersion, 
                    $row.StopProcessInstall, 
                    $row.StopProcessUninstall, 
                    $row.PreScriptInstall, 
                    $row.PostScriptInstall, 
                    $row.PreScriptUninstall, 
                    $row.PostScriptUninstall, 
                    $row.CustomArgumentListInstall, 
                    $row.CustomArgumentListUninstall, 
                    $row.InstallIntent, 
                    $row.Notification, 
                    $row.GroupID
                )
            }
        }
        else {
            $dataGridViewSelected.Rows.Clear()
        }

        # Autosize columns in $dataGridViewSelected after import
        $dataGridViewSelected.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
    })

# Event handler for exporting CSV
$exportButton.Add_Click({
        # Check if DataGridView is not empty
        if ($dataGridViewSelected.Rows.Count -gt 0) {
            # Create an empty array to store the selected data
            $selectedData = @()

            # Iterate through DataGridView rows
            foreach ($row in $dataGridViewSelected.Rows) {
                $packageID = $row.Cells['PackageID'].Value
                $context = $row.Cells['Context'].Value
                $acceptNewerVersion = $row.Cells['AcceptNewerVersion'].Value
                $updateOnly = $row.Cells['UpdateOnly'].Value

                # Check if all required values are not null or empty
                if ($packageID -ne $null -and $packageID -ne '' -and
                    $context -ne $null -and $context -ne '' -and
                    $acceptNewerVersion -ne $null -and $acceptNewerVersion -ne '' -and
                    $updateOnly -ne $null -and $updateOnly -ne '') {
                    # Create a hashtable representing the row data and add it to the selected data array
                    $rowData = [ordered]@{
                        'PackageID'                   = $packageID
                        'Context'                     = $context
                        'AcceptNewerVersion'          = $acceptNewerVersion
                        'UpdateOnly'                  = $updateOnly
                        'TargetVersion'               = $row.Cells['TargetVersion'].Value
                        'StopProcessInstall'          = $row.Cells['StopProcessInstall'].Value
                        'StopProcessUninstall'        = $row.Cells['StopProcessUninstall'].Value
                        'PreScriptInstall'            = $row.Cells['PreScriptInstall'].Value
                        'PostScriptInstall'           = $row.Cells['PostScriptInstall'].Value
                        'PreScriptUninstall'          = $row.Cells['PreScriptUninstall'].Value
                        'PostScriptUninstall'         = $row.Cells['PostScriptUninstall'].Value
                        'CustomArgumentListInstall'   = $row.Cells['CustomArgumentListInstall'].Value
                        'CustomArgumentListUninstall' = $row.Cells['CustomArgumentListUninstall'].Value
                        'InstallIntent'               = $row.Cells['InstallIntent'].Value
                        'Notification'                = $row.Cells['Notification'].Value
                        'GroupID'                     = $row.Cells['GroupID'].Value
                    }
                    $selectedData += New-Object PSObject -Property $rowData
                }
            }

            # Check if any valid data was extracted
            if ($selectedData.Count -gt 0) {
                $defaultFileName = "WinGet-WrapperImportGUI-$timestamp.csv"
                $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveFileDialog.InitialDirectory = (Get-Location).Path
                $saveFileDialog.FileName = $defaultFileName
                $saveFileDialog.Filter = 'CSV files (*.csv)|*.csv'
            
                $result = $saveFileDialog.ShowDialog()
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    $csvFilePath = $saveFileDialog.FileName
                    $selectedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                    Write-ConsoleTextBox "Exported: $csvFilePath"
                }
                else {
                    Write-ConsoleTextBox 'No data exported.'
                }
            }
            else {
                Write-ConsoleTextBox 'No valid data to export.'
            }
        }
        else {
            Write-ConsoleTextBox 'No data to export.'
        }
    })


# Event handler for moving selected rows
$moveButton.Add_Click({
        $selectedRows = $dataGridView.SelectedRows
        foreach ($row in $selectedRows) {
            $name = $row.Cells['Name'].Value
            $id = $row.Cells['ID'].Value
            $version = $row.Cells['Version'].Value
        
            # Add a new row to $dataGridViewSelected
            $rowIndex = $dataGridViewSelected.Rows.Add()
        
            # Set the "PackageID" column with the value from the selected row's "Name" column
            $dataGridViewSelected.Rows[$rowIndex].Cells['PackageID'].Value = $id

            # Set default values for other columns
            $dataGridViewSelected.Rows[$rowIndex].Cells['Context'].Value = 'Machine'
            $dataGridViewSelected.Rows[$rowIndex].Cells['AcceptNewerVersion'].Value = '1'
            $dataGridViewSelected.Rows[$rowIndex].Cells['UpdateOnly'].Value = '0'
        
            # Autosize columns in $dataGridViewSelected after adding rows and setting values
            $dataGridViewSelected.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)

            Write-ConsoleTextBox "Added '$id' to import list"

            # Optionally remove the row from the original DataGridView
            #$dataGridView.Rows.Remove($row)  # Do not remove row after copy to selected datagridview
        }
    })

# Function to write Column Definitions to ConsoleWriteBox
Function GetColumnDefinitions {
    # Fetch the content of the README.md file
    $url = 'https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/README.md'
    Write-ConsoleTextBox "$url"
    $response = Invoke-WebRequest -Uri $url

    # Check if the request was successful
    if ($response.StatusCode -eq 200) {
        $content = $response.Content

        # Get column names dynamically from the DataGridView
        $columnNames = @()
        foreach ($column in $dataGridViewSelected.Columns) {
            $columnNames += $column.Name  # Collect column names into an array
        }
        Write-ConsoleTextBox 'Column Name --> Description'
        Write-ConsoleTextBox '****************************'        
        # Loop through each column name
        foreach ($columnName in $columnNames) {
            # Use regex to find the line corresponding to the column name
            if ($content -match "\* $columnName\s*=\s*(.*?)<br>") {
                $description = $matches[1] -replace '<br>', "`n" -replace '\* ', '' # Clean up the description
                Write-ConsoleTextBox "$columnName --> $description"
            }
            else {
                Write-ConsoleTextBox "$columnName --> No description found."
            }
        }
        Write-ConsoleTextBox '****************************'
    }
    else {
        Write-ConsoleTextBox "Failed to retrieve the README file. Status code: $($response.StatusCode)"
    }
}






# Function to get WinGet package details for a single/selected package (to find possible winget values, system, versions, etc etc.)
function WinGetPackageDetails {
    param (
        [string]$PackageID
    )

    # Get package details and available versions
    $WingetPackageDetails = winget show --id $PackageID --source WinGet --accept-source-agreements --disable-interactivity

    # Output package details line by line to maintain formatting
    Write-ConsoleTextBox "$PackageID - Details:"
    $WingetPackageDetails -split "`r?`n" | ForEach-Object { Write-ConsoleTextBox $_ }
    Write-ConsoleTextBox '_'  # Separator for readability


    # Optionally return the available versions array for further processing
    return $WinGetPackageDetails
}

# Function to get WinGet package versions for a single/selected package (to find possible winget values, system, versions, etc etc.)
function WinGetPackageVersions {
    param (
        [string]$PackageID
    )
    # Get package details and available versions
    $WingetPackageVersionsOutput = winget show --id $PackageID --source WinGet --versions

    # Output available versions header
    Write-ConsoleTextBox "$PackageID - Available Versions:"

    # Initialize an array for the available versions
    $WinGetPackageVersions = @()  

    # Split version details into lines and filter out empty lines
    $versionLines = $WingetPackageVersionsOutput -split "`r?`n" | Where-Object { 
        -not [string]::IsNullOrWhiteSpace($_) 
    }

    # Skip the first three lines and process the remaining lines
    foreach ($line in $versionLines[3..($versionLines.Length - 1)]) {
        $trimmedLine = $line.Trim()  # Trim whitespace
        if (-not [string]::IsNullOrWhiteSpace($trimmedLine) -and $trimmedLine -notmatch '^-+$') {
            # Check it's not empty or dashes
            $WinGetPackageVersions += $trimmedLine  # Add to available versions array
            Write-ConsoleTextBox $trimmedLine  # Display each version line
        }
    }

    # Return the available versions array for further processing
    return $WinGetPackageVersions  # Return the array using the specified variable name
}


# Define a function to parse the search results
function ParseSearchResults($searchResult) {
    Write-ConsoleTextBox 'Parsing data...'
    $parsedData = @()
    $pattern = '^(.+?)\s+((?:[\w.-]+(?:\.[\w.-]+)+))\s+(\S.*?)\s*$'
    $searchResult -split "`n" | Where-Object { $_ -match $pattern } | ForEach-Object {
        $parsedName = $Matches[1].Trim()
        $parsedID = $Matches[2].Trim()
        $parsedID = $parsedID -replace 'ÔÇª', ''  # Remove ellipsis character from ID
        $parsedVersion = $Matches[3].Trim()

        # Add parsed and cleaned data to the result
        $parsedData += [PSCustomObject]@{
            'Name'    = $parsedName
            'ID'      = $parsedID
            'Version' = $parsedVersion
        }
    }
    Write-ConsoleTextBox 'Finished'
    return $parsedData

}
# Define the PerformSearch function that uses the parsing function
function PerformSearch {
    $searchString = $searchBox.Text
    $searchErrorLabel.Text = "Searching for '$searchString'"

    #Update winget sources (to prevent source updating as part of results, wierd output)
    @(winget source update)

    # Search Logic
    if (![string]::IsNullOrWhiteSpace($searchString)) {
        Write-ConsoleTextBox "winget search --query $searchString --source WinGet --accept-source-agreements --disable-interactivity"
        $searchResult = @(winget search --query $searchString --source WinGet --accept-source-agreements --disable-interactivity)
        # Splitting the search result into lines for logging purposes
        $lines = $searchResult -split "`r`n"

        # Writing each line to the consoleTextBox
        foreach ($line in $lines) {
            Write-ConsoleTextBox $line
        }
        if ($searchResult -contains 'No package found matching input criteria.') {
            $dataGridView.Rows.Clear()
            $searchErrorLabel.Text = "No WinGet package found for search query '$searchString'"
        }
        else {
            # Clear existing rows from DataGridView
            $dataGridView.Rows.Clear()

            # Parse the search result using the ParseSearchResults function
            $parsedSearchResult = ParseSearchResults -searchResult $searchResult |
            Where-Object { $null -ne $_.Name -and $_.Name -ne '' -and $_.Name.Trim() -ne '' }

            # Add parsed data to DataGridView
            $parsedSearchResult | ForEach-Object {
                $row = $dataGridView.Rows.Add($_.Name, $_.ID, $_.Version)
            }
            $searchErrorLabel.Text = "Found $($dataGridView.RowCount) packages for '$searchString'"
        }
    }
    else {
        $dataGridView.Rows.Clear()
        $searchErrorLabel.Text = 'Please enter a search query.'
    }
}


# Assign search function to the button click event
$searchButton.Add_Click({ PerformSearch })

# Allow pressing Enter to trigger search
$form.KeyPreview = $true
$form.Add_KeyDown({
        param($keySender, $keyEvent)
        if ($keyEvent.KeyCode -eq 'Enter') {
            PerformSearch
        }
    })

# Event Handler for Get package details button
$GetPackageDetails.Add_Click({
        $selectedRows = $dataGridView.SelectedRows
        foreach ($row in $selectedRows) {
            $id = $row.Cells['ID'].Value
            WinGetPackageDetails -PackageID "$id"
        }
    })

# Event Handler for Get package versions button
$GetPackageVersions.Add_Click({
        $selectedRows = $dataGridView.SelectedRows
        foreach ($row in $selectedRows) {
            $id = $row.Cells['ID'].Value
            WinGetPackageVersions -PackageID "$id"
        }
    })

# Event handler for the "Import to InTune" button
$InTuneimportButton.Add_Click({
        Write-ConsoleTextBox 'Started import to InTune..'

        # Check if $tenantIDTextBox.Text is empty, matches $tenantIDTextBoxDefaultText, or does not contain a dot
        if ([string]::IsNullOrWhiteSpace($tenantIDTextBox.Text) -or $tenantIDTextBox.Text -eq $tenantIDTextBoxDefaultText -or -not ($tenantIDTextBox.Text -like '*.*')) {
            Write-ConsoleTextBox 'Please enter a valid Tenant ID before importing to InTune.'
            return  # Stop further execution
        }

        # List of files to check
        $filesToCheck = @(
            'WinGet-Wrapper.ps1',
            'WinGet-WrapperDetection.ps1',
            'WinGet-WrapperRequirements.ps1',
            'WinGet-WrapperImportFromCSV.ps1',
            'IntuneWinAppUtil.exe'
        )

        $foundAllFiles = $true
        foreach ($file in $filesToCheck) {
            $fileFullPath = Join-Path -Path $scriptRoot -ChildPath $file

            if (-not (Test-Path -Path $fileFullPath -PathType Leaf)) {
                # File not found, write a message to the console text box
                Write-ConsoleTextBox "File '$file' was not found."
                $foundAllFiles = $false
            }
            else {
                # File found, write a message to the console text box
                Write-ConsoleTextBox "File '$file' was found."
            }
        }

        if ($foundAllFiles) {
            Write-ConsoleTextBox 'All required files found. Continue import to InTune...'

            # Export DataGridViewSelected to CSV - Save CSV Temporary
            $selectedData = @()

            # Check if DataGridView is not empty
            if ($dataGridViewSelected.Rows.Count -gt 0) {
                # Create an empty array to store the selected data
                $selectedData = @()

                # Iterate through DataGridView rows
                foreach ($row in $dataGridViewSelected.Rows) {
                    $packageID = $row.Cells['PackageID'].Value
                    $context = $row.Cells['Context'].Value
                    $acceptNewerVersion = $row.Cells['AcceptNewerVersion'].Value
                    $updateOnly = $row.Cells['UpdateOnly'].Value

                    # Check if all required values are not null or empty
                    if ($packageID -ne $null -and $packageID -ne '' -and
                        $context -ne $null -and $context -ne '' -and
                        $acceptNewerVersion -ne $null -and $acceptNewerVersion -ne '' -and
                        $updateOnly -ne $null -and $updateOnly -ne '') {
                        # Create a hashtable representing the row data and add it to the selected data array
                        $rowData = [ordered]@{
                            'PackageID'                   = $packageID
                            'Context'                     = $context
                            'AcceptNewerVersion'          = $acceptNewerVersion
                            'UpdateOnly'                  = $updateOnly
                            'TargetVersion'               = $row.Cells['TargetVersion'].Value
                            'StopProcessInstall'          = $row.Cells['StopProcessInstall'].Value
                            'StopProcessUninstall'        = $row.Cells['StopProcessUninstall'].Value
                            'PreScriptInstall'            = $row.Cells['PreScriptInstall'].Value
                            'PostScriptInstall'           = $row.Cells['PostScriptInstall'].Value
                            'PreScriptUninstall'          = $row.Cells['PreScriptUninstall'].Value
                            'PostScriptUninstall'         = $row.Cells['PostScriptUninstall'].Value
                            'CustomArgumentListInstall'   = $row.Cells['CustomArgumentListInstall'].Value
                            'CustomArgumentListUninstall' = $row.Cells['CustomArgumentListUninstall'].Value
                            'InstallIntent'               = $row.Cells['InstallIntent'].Value
                            'Notification'                = $row.Cells['Notification'].Value
                            'GroupID'                     = $row.Cells['GroupID'].Value
                        }
                        $selectedData += New-Object PSObject -Property $rowData
                    }
                }
            }
            if ($selectedData -ne $null -and $selectedData.Count -gt 0) {
                $fileName = "TempExport-$timestamp.csv"  # Construct the filename with timestamp
                $csvFilePath = Join-Path -Path $scriptRoot -ChildPath $fileName  # Save to script root directory
                $selectedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                Write-ConsoleTextBox "Exported: $csvFilePath"
            }
            else {
                Write-ConsoleTextBox 'No data to export.'
                return  # Stop further execution
            }

            # Prepare the Import script.
            $logFile = "$scriptRoot\Logs\WinGet_WrapperImportFromCSV_$($TimeStamp).log"
            $importScriptPath = Join-Path -Path $scriptRoot -ChildPath 'Winget-WrapperImportFromCSV.ps1'
            Write-ConsoleTextBox "ImportScriptPath: $importScriptPath"

            #Inform user log file location:
            Write-ConsoleTextBox '****************************************************'
            Write-ConsoleTextBox "See log file for progress: $logFile"
            Write-ConsoleTextBox '****************************************************'

            #Run The Import Script
            # Define the arguments to be passed to the script
            $arguments = "-csvFile `"$csvFilePath`" -TenantID $($tenantIDTextBox.Text) -ClientID $($clientIDTextBox.Text) -RedirectURL `"$($redirectURLTextBox.Text)`" -ScriptRoot `"$scriptRoot`" -SkipConfirmation -SkipModuleCheck"
            Write-ConsoleTextBox "Arguments to be passed: $arguments"
            #Set-ExecutionPolicy Bypass -Scope Process -ExecutionPolicy Bypass -Force
            Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy Bypass', "-File `"$importScriptPath`"", $arguments -Wait -NoNewWindow

            # Run Update-GUIFromLogFile in the main thread
            Start-Sleep -Seconds 5  # wait log file creation before reading it
            Update-GUIFromLogFile -logFilePath "$logFile"

            # Remove TempExport-$timestamp.csv
            if (Test-Path $csvFilePath) {
                Remove-Item $csvFilePath -Force
                Write-ConsoleTextBox "File $csvFilePath deleted successfully."
            }
            else {
                Write-ConsoleTextBox "File $csvFilePath not found."
            }

            Write-ConsoleTextBox '****************************************************'
            Write-ConsoleTextBox "Import Log File: $logFile"
            Write-ConsoleTextBox '****************************************************'
        }
        else {
            Write-ConsoleTextBox 'Not all required files were found. Code will not run.'
        }
    })

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 100 -Status 'Sucessfully loaded Winget-Wrapper Import GUI'

# Greeting
Write-ConsoleTextBox '****************************************************'
Write-ConsoleTextBox '                           WinGet-Wrapper'
Write-ConsoleTextBox '  https://github.com/SorenLundt/WinGet-Wrapper'
Write-ConsoleTextBox ''
Write-ConsoleTextBox '              GNU General Public License v3'
Write-ConsoleTextBox '****************************************************'

# Show form
$form.ShowDialog() | Out-Null