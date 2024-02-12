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

# Greeting´
write-host ""
Write-Host "****************************************************"
Write-Host "                  WinGet-Wrapper"
Write-Host "  https://github.com/SorenLundt/WinGet-Wrapper"
Write-Host ""
Write-Host "          GNU General Public License v3"
Write-Host "****************************************************"
write-host ""

Write-Host "Loading Winget-WrapperImportGUI.."

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set the timestamp for log file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

#Find Script root path  
if (-not $PSScriptRoot) {
    $scriptRoot = (Get-Location -PSProvider FileSystem).ProviderPath
} else {
    $scriptRoot = $PSScriptRoot
}

# Create logs folder if it doesn't exist
$LogFolder = Join-Path -Path $scriptRoot -ChildPath "Logs"
    # Create logs folder if it doesn't exist
    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory | Out-Null
    }

# Install and load required modules
$intuneWin32AppModule = "IntuneWin32App"
$microsoftGraphIntuneModule = "Microsoft.Graph.Intune"

#DEBUG (Skip ModuleCheck)
#$SKIPMODULECHECK = $true
if (-not $SKIPMODULECHECK) {

Write-Host "Checking and updating required modules.."

# Check IntuneWin32App module
Write-Host "Checking module $intuneWin32AppModule.."
$moduleInstalled = Get-InstalledModule -Name $intuneWin32AppModule -ErrorAction SilentlyContinue
if (-not $moduleInstalled) {
    Install-Module -Name $intuneWin32AppModule -Force
} else {
    $latestVersion = (Find-Module -Name $intuneWin32AppModule).Version
    if ($moduleInstalled.Version -lt $latestVersion) {
        Update-Module -Name $intuneWin32AppModule -Force
    } else {
        Write-Host "Module $intuneWin32AppModule is already up-to-date." -ForegroundColor Green
    }
}

# Check Microsoft.Graph.Intune module
Write-Host "Checking module $microsoftGraphIntuneModule.."
$moduleInstalled = Get-InstalledModule -Name $microsoftGraphIntuneModule -ErrorAction SilentlyContinue

if (-not $moduleInstalled) {
    Install-Module -Name $microsoftGraphIntuneModule -Force
} else {
    $latestVersion = (Find-Module -Name $microsoftGraphIntuneModule).Version
    if ($moduleInstalled.Version -lt $latestVersion) {
        Update-Module -Name $microsoftGraphIntuneModule -Force
    } else {
        Write-Host "Module $microsoftGraphIntuneModule is already up-to-date." -ForegroundColor Green
    }
}
}
#Import modules
Write-Host "Importing module $intuneWin32AppModule.."
Import-Module -Name "IntuneWin32App"
Write-Host "Importing module $microsoftGraphIntuneModule.."
Import-Module -Name "Microsoft.Graph.Intune"

#Functions
function Write-ConsoleTextBox {
    param (
        [string]$Message,
        [switch]$NoTimeStamp
    )

    if (-not $NoTimeStamp) {
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
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

Write-Host "Loading GUI elements.."

# GUI
# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Winget-Wrapper Import GUI  - https://github.com/SorenLundt/WinGet-Wrapper"
$form.Width = 1475
$form.Height = 900
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.Topmost = $false
$form.MaximizeBox = $False

# Set the icon for the form
write-host "$scriptRoot"
$iconPath = Join-Path -Path $scriptRoot -ChildPath "Winget-Wrapper.ico"
if (Test-Path $iconPath) {
    $icon = New-Object System.Drawing.Icon($iconPath)
    $form.Icon = $icon
} else {
    Write-Host "Icon file not found at $iconPath"
}

# Create TextBox for search string
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10, 10)
$searchBox.Width = 400
$form.Controls.Add($searchBox)

# Create Button for search
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(420, 8)
$searchButton.Width = 50
$searchButton.Text = "Search"
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
$resultsLabel.Text = "WinGet Packages Search Results"
$form.Controls.Add($resultsLabel)

# Create DataGridView for results
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 60)
$dataGridView.Width = 600
$dataGridView.Height = 500
$form.Controls.Add($dataGridView)

# Add columns to DataGridView
$dataGridView.Columns.Add("Name", "Name")
$dataGridView.Columns.Add("ID", "ID")
$dataGridView.Columns.Add("Version", "Version")

# Set initial widths for columns in the DataGridView
$dataGridView.Columns['Name'].Width = 200  # Adjust the width as needed
$dataGridView.Columns['ID'].Width = 150    # Adjust the width as needed
$dataGridView.Columns['Version'].Width = 100  # Adjust the width as needed

# Create Label for $dataGridView (Selected)
$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Location = New-Object System.Drawing.Point(650, 37)
$resultsLabel.Width = 200
$resultsLabel.Text = "Selected WinGet Packages"
$form.Controls.Add($resultsLabel)

# Create a second DataGridView
$dataGridViewSelected = New-Object System.Windows.Forms.DataGridView
$dataGridViewSelected.Location = New-Object System.Drawing.Point(650, 60)
$dataGridViewSelected.Width = 800
$dataGridViewSelected.Height = 500
$form.Controls.Add($dataGridViewSelected)

# Add columns to the second DataGridView
$dataGridViewSelected.Columns.Add("PackageID", "PackageID")
$dataGridViewSelected.Columns.Add("Context", "Context")
$dataGridViewSelected.Columns.Add("AcceptNewerVersion", "AcceptNewerVersion")
$dataGridViewSelected.Columns.Add("UpdateOnly", "UpdateOnly")
$dataGridViewSelected.Columns.Add("TargetVersion", "TargetVersion")
$dataGridViewSelected.Columns.Add("StopProcessInstall", "StopProcessInstall")
$dataGridViewSelected.Columns.Add("StopProcessUninstall", "StopProcessUninstall")
$dataGridViewSelected.Columns.Add("PreScriptInstall", "PreScriptInstall")
$dataGridViewSelected.Columns.Add("PostScriptInstall", "PostScriptInstall")
$dataGridViewSelected.Columns.Add("PreScriptUninstall", "PreScriptUninstall")
$dataGridViewSelected.Columns.Add("PostScriptUninstall", "PostScriptUninstall")
$dataGridViewSelected.Columns.Add("CustomArgumentListInstall", "CustomArgumentListInstall")
$dataGridViewSelected.Columns.Add("CustomArgumentListUninstall", "CustomArgumentListUninstall")
$dataGridViewSelected.Columns.Add("InstallIntent", "InstallIntent")
$dataGridViewSelected.Columns.Add("Notification", "Notification")
$dataGridViewSelected.Columns.Add("GroupID", "GroupID")

# Set initial widths for columns in the DataGridViewSelected
foreach ($column in $dataGridViewSelected.Columns){
    $column.Width = "80"
}

# Create Button for exporting CSV from dataGridViewSelected
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(750, 560)
$exportButton.Width = 100
$exportButton.Text = "Export CSV"
$form.Controls.Add($exportButton)

# Create Button for moving selected rows with a right-arrow icon
$moveButton = New-Object System.Windows.Forms.Button
$moveButton.Location = New-Object System.Drawing.Point(618, 280)
$moveButton.Width = 30
$moveButton.Height = 30

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
$moveButton.Text = ""
$form.Controls.Add($moveButton)

# Create Button for deleting selected rows
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Location = New-Object System.Drawing.Point(1350, 560)
$deleteButton.Width = 100
$deleteButton.Text = "Delete Selected"
$form.Controls.Add($deleteButton)

# Create Button for importing CSV to dataGridViewSelected
$importCSVButton = New-Object System.Windows.Forms.Button
$importCSVButton.Location = New-Object System.Drawing.Point(650, 560)
$importCSVButton.Width = 100
$importCSVButton.Text = "Import CSV"
$form.Controls.Add($importCSVButton)

# Create a TextBox for console output
$consoleTextBox = New-Object System.Windows.Forms.TextBox
$consoleTextBox.Location = New-Object System.Drawing.Point(10, 650)
$consoleTextBox.Width = 1420
$consoleTextBox.Height = 200
$consoleTextBox.Multiline = $true
$consoleTextBox.ScrollBars = "Vertical"
$form.Controls.Add($consoleTextBox)

# Create Button for importing data into InTune
$InTuneimportButton = New-Object System.Windows.Forms.Button
$InTuneimportButton.Location = New-Object System.Drawing.Point(650, 610)
$InTuneimportButton.Width = 120
$InTuneimportButton.Text = "Import to InTune"
$IntuneImportButton.Visible = $True
$form.Controls.Add($InTuneimportButton)

# Create TextBox for Tenant ID
$tenantIDTextBox = New-Object System.Windows.Forms.TextBox
$tenantIDTextBox.Location = New-Object System.Drawing.Point(650, 585)
$tenantIDTextBox.Width = 300
$form.Controls.Add($tenantIDTextBox)

# Help text for Tenant ID textbox
$tenantIDTextBoxDefaultText = "Enter Tenant ID (e.g., company.onmicrosoft.com)"
$tenantIDTextBox.Text = "$tenantIDTextBoxDefaultText"
$tenantIDTextBox.ForeColor = [System.Drawing.Color]::Gray

#OutPut to textbox in GUI

Write-Host "Loading Event handlers"
# EVENTS
# Event handler for when the textbox gains focus (Enter event)
$tenantIDTextBox.Add_Enter({
    if ($tenantIDTextBox.Text -eq "$tenantIDTextBoxDefaultText") {
        $tenantIDTextBox.Text = ""
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

# Event handler for deleting selected rows
$deleteButton.Add_Click({
    foreach ($row in $dataGridViewSelected.SelectedRows) {
        $dataGridViewSelected.Rows.Remove($row)
    }

    # Autosize columns after deleting rows
    $dataGridViewSelected.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
})
# Event handler for importing CSV
$importCSVButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = (Get-Location).Path
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    
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
    } else {
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
                    'PackageID' = $packageID
                    'Context' = $context
                    'AcceptNewerVersion' = $acceptNewerVersion
                    'UpdateOnly' = $updateOnly
                    'TargetVersion' = $row.Cells['TargetVersion'].Value
                    'StopProcessInstall' = $row.Cells['StopProcessInstall'].Value
                    'StopProcessUninstall' = $row.Cells['StopProcessUninstall'].Value
                    'PreScriptInstall' = $row.Cells['PreScriptInstall'].Value
                    'PostScriptInstall' = $row.Cells['PostScriptInstall'].Value
                    'PreScriptUninstall' = $row.Cells['PreScriptUninstall'].Value
                    'PostScriptUninstall' = $row.Cells['PostScriptUninstall'].Value
                    'CustomArgumentListInstall' = $row.Cells['CustomArgumentListInstall'].Value
                    'CustomArgumentListUninstall' = $row.Cells['CustomArgumentListUninstall'].Value
                    'InstallIntent' = $row.Cells['InstallIntent'].Value
                    'Notification' = $row.Cells['Notification'].Value
                    'GroupID' = $row.Cells['GroupID'].Value
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
            $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
            
            $result = $saveFileDialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $csvFilePath = $saveFileDialog.FileName
                $selectedData | Export-Csv -Path $csvFilePath -NoTypeInformation
                Write-ConsoleTextBox "Exported: $csvFilePath"
            } else {
                Write-ConsoleTextBox "No data exported."
            }
        } else {
            Write-ConsoleTextBox "No valid data to export."
        }
    } else {
        Write-ConsoleTextBox "No data to export."
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
        $dataGridViewSelected.Rows[$rowIndex].Cells['Context'].Value = "Machine"
        $dataGridViewSelected.Rows[$rowIndex].Cells['AcceptNewerVersion'].Value = "1"
        $dataGridViewSelected.Rows[$rowIndex].Cells['UpdateOnly'].Value = "1"
        
        # Autosize columns in $dataGridViewSelected after adding rows and setting values
    $dataGridViewSelected.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)

        # Optionally remove the row from the original DataGridView
        #$dataGridView.Rows.Remove($row)  # Do not remove row after copy to selected datagridview
    }
})

# Define a function to parse the search results
function ParseSearchResults($searchResult) {
    Write-ConsoleTextBox "Parsing data..."
    $parsedData = @()
    $pattern = "^(.+?)\s+((?:[\w.-]+(?:\.[\w.-]+)+))\s+(\S.*?)\s*$"
    $searchResult -split "`n" | Where-Object { $_ -match $pattern } | ForEach-Object {
        $parsedName = $Matches[1].Trim()
        $parsedID = $Matches[2].Trim()
        $parsedID = $parsedID -replace 'ÔÇª', ''  # Remove ellipsis character from ID
        $parsedVersion = $Matches[3].Trim()

        # Add parsed and cleaned data to the result
        $parsedData += [PSCustomObject]@{
            'Name' = $parsedName
            'ID' = $parsedID
            'Version' = $parsedVersion
        }
    }
    Write-ConsoleTextBox "Finished"
    return $parsedData

}
# Define the PerformSearch function that uses the parsing function
function PerformSearch {
    #Clear consoleTextBox before running code
    $consoleTextBox.Clear()
    $searchString = $searchBox.Text
    $searchErrorLabel.Text = "Searching for '$searchString'"

    # Your existing search logic here
    if (![string]::IsNullOrWhiteSpace($searchString)) {
        Write-ConsoleTextBox "winget search --query $searchString --source WinGet --accept-source-agreements --disable-interactivity"
        $searchResult = @(winget search --query $searchString --source WinGet --accept-source-agreements --disable-interactivity)
    # Splitting the search result into lines for logging purposes
    $lines = $searchResult -split "`r`n"

    # Writing each line to the consoleTextBox
    foreach ($line in $lines) {
        Write-ConsoleTextBox $line
    }
        if ($searchResult -contains "No package found matching input criteria.") {
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
        $searchErrorLabel.Text = "Finished searching for '$searchString'"
        }
    } else {
        $dataGridView.Rows.Clear()
        $searchErrorLabel.Text = "Please enter a search query."
    }
}


# Assign search function to the button click event
$searchButton.Add_Click({ PerformSearch })

# Allow pressing Enter to trigger search
$form.KeyPreview = $true
$form.Add_KeyDown({
    param($keySender, $keyEvent)
    if ($keyEvent.KeyCode -eq "Enter") {
        PerformSearch
    }
})

# Event handler for the "Import to InTune" button
$InTuneimportButton.Add_Click({
    # Clear consoleTextBox before running code
    $consoleTextBox.Clear()

    Write-ConsoleTextBox "Started import to InTune.."

    # Check if $tenantIDTextBox.Text is empty, matches $tenantIDTextBoxDefaultText, or does not contain a dot
    if ([string]::IsNullOrWhiteSpace($tenantIDTextBox.Text) -or $tenantIDTextBox.Text -eq $tenantIDTextBoxDefaultText -or -not ($tenantIDTextBox.Text -like '*.*')) {
        Write-ConsoleTextBox "Please enter a valid Tenant ID before importing to InTune."
        return  # Stop further execution
    }

    # List of files to check
    $filesToCheck = @(
        "WinGet-Wrapper.ps1",
        "WinGet-WrapperDetection.ps1",
        "WinGet-WrapperRequirements.ps1",
        "WinGet-WrapperImportFromCSV.ps1",
        "IntuneWinAppUtil.exe"
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
        Write-ConsoleTextBox "All required files found. Continue import to InTune..."

        # Export DataGridViewSelected to CSV - Save CSV Temporary
        $selectedData = @()

        # Check if DataGridView is not empty
        if ($dataGridViewSelected.Rows.Count -gt 0) {
            # Get the row data if DataGridView is not empty
            $row = $dataGridViewSelected.Rows[0]  # Assuming there's only one row
            $packageID = $row.Cells['PackageID'].Value
            $context = $row.Cells['Context'].Value
            $acceptNewerVersion = $row.Cells['AcceptNewerVersion'].Value
            $updateOnly = $row.Cells['UpdateOnly'].Value
            
            if ($packageID -ne $null -and $packageID -ne '' -and
                $context -ne $null -and $context -ne '' -and
                $acceptNewerVersion -ne $null -and $acceptNewerVersion -ne '' -and
                $updateOnly -ne $null -and $updateOnly -ne '') {
                [PSCustomObject]@{
                    'PackageID' = $packageID
                    'Context' = $context
                    'AcceptNewerVersion' = $acceptNewerVersion
                    'UpdateOnly' = $updateOnly
                    'TargetVersion' = $row.Cells['TargetVersion'].Value
                    'StopProcessInstall' = $row.Cells['StopProcessInstall'].Value
                    'StopProcessUninstall' = $row.Cells['StopProcessUninstall'].Value
                    'PreScriptInstall' = $row.Cells['PreScriptInstall'].Value
                    'PostScriptInstall' = $row.Cells['PostScriptInstall'].Value
                    'PreScriptUninstall' = $row.Cells['PreScriptUninstall'].Value
                    'PostScriptUninstall' = $row.Cells['PostScriptUninstall'].Value
                    'CustomArgumentListInstall' = $row.Cells['CustomArgumentListInstall'].Value
                    'CustomArgumentListUninstall' = $row.Cells['CustomArgumentListUninstall'].Value
                    'InstallIntent' = $row.Cells['InstallIntent'].Value
                    'Notification' = $row.Cells['Notification'].Value
                    'GroupID' = $row.Cells['GroupID'].Value
                } | ForEach-Object { $selectedData += $_ }
            }
        }

        if ($selectedData -ne $null -and $selectedData.Count -gt 0) {
            $fileName = "TempExport-$timestamp.csv"  # Construct the filename with timestamp
            $csvFilePath = Join-Path -Path $scriptRoot -ChildPath $fileName  # Save to script root directory
            $selectedData | Export-Csv -Path $csvFilePath -NoTypeInformation
            Write-ConsoleTextBox "Exported: $csvFilePath"
        } else {
            Write-ConsoleTextBox "No data to export."
            return  # Stop further execution
        }

        # Prepare the Import script.
        $logFile = "$scriptRoot\Logs\WinGet_WrapperImportFromCSV_$($TimeStamp).log"
        $importScriptPath = Join-Path -Path $scriptRoot -ChildPath "Winget-WrapperImportFromCSV.ps1"
        Write-ConsoleTextBox "ImportScriptPath: $importScriptPath"

        #Inform user log file location:
        Write-ConsoleTextBox "****************************************************"
        Write-ConsoleTextBox "See log file for progress: $logFile"
        Write-ConsoleTextBox "****************************************************"

        #Run The Import Script
    # Define the arguments to be passed to the script
    $arguments = "-csvFile `"$csvFilePath`" -TenantID $($tenantIDTextBox.Text) -LogFile `"$logFile`" -ScriptRoot `"$scriptRoot`" -SkipConfirmation -SkipModuleCheck"
    Write-ConsoleTextBox "Arguments to be passed: $arguments"
    Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File `"$importScriptPath`"", $arguments -Wait -NoNewWindow

        # Run Update-GUIFromLogFile in the main thread
        Start-Sleep -Seconds 5  # wait log file creation before reading it
        Update-GUIFromLogFile -logFilePath "$logFile"

        # Remove TempExport-$timestamp.csv
        if (Test-Path $csvFilePath) {
            Remove-Item $csvFilePath -Force
            Write-ConsoleTextBox "File $csvFilePath deleted successfully."
        } else {
            Write-ConsoleTextBox "File $csvFilePath not found."
        }

        Write-ConsoleTextBox "****************************************************"
        Write-ConsoleTextBox "Import Log File: $logFile"
        Write-ConsoleTextBox "****************************************************"
    } else {
        Write-ConsoleTextBox "Not all required files were found. Code will not run."
    }
})

# Greeting
Write-ConsoleTextBox "****************************************************"
Write-ConsoleTextBox "                           WinGet-Wrapper"
Write-ConsoleTextBox "  https://github.com/SorenLundt/WinGet-Wrapper"
Write-ConsoleTextBox ""
Write-ConsoleTextBox "              GNU General Public License v3"
Write-ConsoleTextBox "****************************************************"

# Show form
$form.ShowDialog() | Out-Null