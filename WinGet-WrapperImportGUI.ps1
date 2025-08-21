# Soren Lundt - 12-02-2024
# URL: https://github.com/SorenLundt/WinGet-Wrapper
# License: https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/LICENSE.txt
# Graphical interface for WinGet-Wrapper Import - WPF Version (Optimized)
# Package content is stored under Packages\Package.ID-Context-UpdateOnly-UserName-yyyy-mm-dd-hhssmm

# Requirements:
# Requires Script files and IntuneWinAppUtil.exe to be present in script directory
#
# Version History
# Version 1.0 - 12-02-2024 SorenLundt - Initial Version
# Version 1.1 - 21-02-2024 SorenLundt - Fixed issue where only 1 package was imported to InTune
# Version 1.2 - 30-10-2024 SorenLundt - Various improvements
# Version 1.3 - 31-10-2024 SorenLundt - Added Column Definitions help button
# Version 2.0 - 14-07-2025 - Converted to WPF for modern UI
# Version 2.1 - 14-07-2025 - UX/UI Optimizations and enhancements

#Parameters
Param (
    #Skip module checking (for testing purposes)
    [Switch]$SKIPMODULECHECK = $true
)

# Greeting
Write-Host ""
Write-Host "****************************************************"
Write-Host "                  WinGet-Wrapper)"
Write-Host "  https://github.com/SorenLundt/WinGet-Wrapper"
Write-Host ""
Write-Host "          GNU General Public License v3"
Write-Host "****************************************************"
Write-Host "   WinGet-WrapperImportGUI Starting up.."
Write-Host ""

# Function to show loading progress bar in console.
function Show-ConsoleProgress {
    param (
        [string]$Activity = "Loading Winget-Wrapper Import GUI",
        [string]$Status = "",
        [int]$PercentComplete = 0
    )
    Write-Host "$Status - [$PercentComplete%]"
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Update ConsoleProgress
Show-ConsoleProgress -PercentComplete 0 -Status "Initializing..."

# Add required assemblies for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Set the timestamp for log file
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

#Find Script root path  
if (-not $PSScriptRoot) {
    $scriptRoot = (Get-Location -PSProvider FileSystem).ProviderPath
}
else {
    $scriptRoot = $PSScriptRoot
}

# Create logs folder if it doesn't exist
$LogFolder = Join-Path -Path $scriptRoot -ChildPath "Logs"
if (-not (Test-Path -Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory | Out-Null
}

# Install and load required modules
$intuneWin32AppModule = "IntuneWin32App"
$microsoftGraphAuthenticationModule = "Microsoft.Graph.Authentication"
$wingetClientModule = "Microsoft.WinGet.Client"

if (-not $SKIPMODULECHECK) {
    # Check IntuneWin32App module
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

    # Check Microsoft.Graph.Authentication module (still needed for some operations)
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

    # Check Microsoft.WinGet.Client module
    Show-ConsoleProgress -PercentComplete 60 -Status "Checking and updating $wingetClientModule.."
    $moduleInstalled = Get-InstalledModule -Name $wingetClientModule -ErrorAction SilentlyContinue
    if (-not $moduleInstalled) {
        Install-Module -Name $wingetClientModule -Force
    }
    else {
        $latestVersion = (Find-Module -Name $wingetClientModule).Version
        if ($moduleInstalled.Version -lt $latestVersion) {
            Update-Module -Name $wingetClientModule -Force
        }
        else {
            Write-Host "Module $wingetClientModule is already up-to-date." -ForegroundColor Green
        }
    }
}

#Import modules
Show-ConsoleProgress -PercentComplete 70 -Status "Importing module $intuneWin32AppModule.."
Import-Module -Name "IntuneWin32App"
Import-Module -Name "Microsoft.WinGet.Client"

Show-ConsoleProgress -PercentComplete 90 -Status "Unblocking script files (Unblock-File)"
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
Show-ConsoleProgress -PercentComplete 80 -Status "Loading functions.."   

# Define optimized WPF XAML with enhanced UX/UI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"
        Title="WinGet-Wrapper Import GUI v2.1 - Enhanced UX" 
        Height="1100" Width="1600"
        Background="#F5F7FA" 
        FontFamily="Segoe UI" 
        FontSize="12"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        MinHeight="800" MinWidth="1200">
    
    <Window.Resources>
        <!-- Enhanced Button Styles -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect Color="#000000" Opacity="0.15" BlurRadius="8" ShadowDepth="2"/>
                </Setter.Value>
            </Setter>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="6"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#CCCCCC"/>
                                <Setter Property="Foreground" Value="#666666"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#6C757D"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#5A6268"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#495057"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#DC3545"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#C82333"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#BD2130"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#28A745"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#218838"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#1E7E34"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Enhanced TextBox Style with Validation Support -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderBrush" Value="#DDE2E8"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Height" Value="42"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Height="{TemplateBinding Height}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="12,0,12,0" VerticalAlignment="Center" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Hidden"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="BorderBrush" Value="#80BDFF"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#0078D7"/>
                                <Setter Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect Color="#0078D7" Opacity="0.3" BlurRadius="8" ShadowDepth="0"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Enhanced DataGrid Style -->
        <Style x:Key="ModernDataGrid" TargetType="DataGrid">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E1E5E9"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#F8F9FA"/>
            <Setter Property="VerticalGridLinesBrush" Value="Transparent"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#FAFBFC"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="CanUserResizeColumns" Value="True"/>
            <Setter Property="CanUserSortColumns" Value="True"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="RowHeight" Value="36"/>
        </Style>

        <!-- Enhanced DataGrid Header Style -->
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#495057"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderBrush" Value="#495057"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <!-- Enhanced Card Style -->
        <Style x:Key="SectionCard" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E1E5E9"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
            <Setter Property="Margin" Value="8"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect Color="#000000" Opacity="0.08" BlurRadius="12" ShadowDepth="3"/>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Section Header Style -->
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#2C3E50"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>

        <!-- Icon Button Style for Transfer -->
        <Style x:Key="TransferButton" TargetType="Button">
            <Setter Property="Background" Value="#28A745"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Width" Value="60"/>
            <Setter Property="Height" Value="60"/>
            <Setter Property="FontSize" Value="20"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="30"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#218838"/>
                                <Setter Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect Color="#28A745" Opacity="0.4" BlurRadius="15" ShadowDepth="3"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#1E7E34"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Validation Icon Style -->
        <Style x:Key="ValidationIcon" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Right"/>
            <Setter Property="Margin" Value="8,0"/>
        </Style>
    </Window.Resources>

    <!-- Main Grid Layout -->
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="2*"/>      <!-- Main Content Area -->
            <RowDefinition Height="1*"/>      <!-- Console Output -->
        </Grid.RowDefinitions>

        <!-- MAIN CONTENT AREA -->
        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>   <!-- Left Column: WinGet Packages -->
                <ColumnDefinition Width="3*"/>   <!-- Right Column: Intune -->
            </Grid.ColumnDefinitions>

            <!-- LEFT COLUMN: WINGET PACKAGES & TRANSFER -->
            <Grid Grid.Column="0" Margin="0,0,8,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- WinGet Packages Panel (Search and Results Merged) -->
                <Border Grid.Row="0" Style="{StaticResource SectionCard}" Margin="0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <TextBlock Grid.Row="0" Text="WinGet Packages" Style="{StaticResource SectionHeader}"/>

                        <!-- Search Input -->
                        <Grid Grid.Row="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox Name="SearchBox" Grid.Column="0" Style="{StaticResource ModernTextBox}"
                                     Text="Search for software (e.g., VLC, Chrome)"
                                     Foreground="Gray"
                                     ToolTip="Enter a software name to search the WinGet repository"/>
                            <Button Name="SearchButton" Grid.Column="1" Content="Search" Style="{StaticResource PrimaryButton}" 
                                    Width="120" Height="38" Margin="12,0,0,0"/>
                        </Grid>

                        <!-- Search Status & Help -->
                        <Grid Grid.Row="2" Margin="0,12,0,0">
                            <TextBlock Name="SearchStatus" FontSize="13" Foreground="#6C757D" FontStyle="Italic" VerticalAlignment="Center"/>
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                <TextBlock Name="GitHubLink" Text="GitHub Repository" 
                                           Foreground="#0078D7" TextDecorations="Underline" 
                                           Cursor="Hand" FontSize="12" Margin="0,0,12,0"/>
                            </StackPanel>
                        </Grid>

                        <!-- Quick Filter -->
                        <TextBox Name="QuickFilter" Grid.Row="3" Style="{StaticResource ModernTextBox}"
                                 Text="Quick filter results..." Foreground="Gray" Margin="0,12,0,12"
                                 ToolTip="Type here to filter the search results below"/>

                        <!-- Enhanced DataGrid with Better Spacing -->
                        <DataGrid Name="SearchResults" Grid.Row="4" Style="{StaticResource ModernDataGrid}">
                            <DataGrid.Columns>
                                <DataGridTemplateColumn Width="40">
                                    <DataGridTemplateColumn.Header>
                                        <CheckBox Name="SelectAllCheckBox" ToolTip="Select/Deselect All"/>
                                    </DataGridTemplateColumn.Header>
                                    <DataGridTemplateColumn.CellTemplate>
                                        <DataTemplate>
                                            <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                                                      IsChecked="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                                        </DataTemplate>
                                    </DataGridTemplateColumn.CellTemplate>
                                </DataGridTemplateColumn>
                                <DataGridTextColumn Header="Package Name" Binding="{Binding Name}" Width="*" MinWidth="180"/>
                                <DataGridTextColumn Header="Package ID" Binding="{Binding ID}" Width="2*"/>
                                <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>

                        <!-- Action Buttons with Clear Grouping -->
                        <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,16,0,0" HorizontalAlignment="Left">
                            <Button Name="GetPackageDetails" Content="Details" 
                                    Style="{StaticResource SecondaryButton}" Width="100"/>
                            <Button Name="GetPackageVersions" Content="Versions" 
                                    Style="{StaticResource SecondaryButton}" Width="110"/>
                            <Button Name="ViewOnGitHub" Content="View on GitHub" 
                                    Style="{StaticResource SecondaryButton}" Width="120"/>
                            <Button Name="TransferButton" Content="Add to Import List" 
                                    Style="{StaticResource SuccessButton}" Width="160" Margin="12,6,6,6"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Enhanced Transfer Area with Visual Appeal -->
                <Grid Grid.Row="1" VerticalAlignment="Center" Margin="0,12,0,0">
                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
                        <TextBlock Name="TransferCount" Text="" FontSize="10" Foreground="#28A745" 
                                   HorizontalAlignment="Center" Margin="0,4,0,0" FontWeight="SemiBold"/>
                    </StackPanel>
                </Grid>
            </Grid>

            <!-- RIGHT COLUMN: INTUNE CONFIG & IMPORT LIST -->
            <Grid Grid.Column="1" Margin="8,0,0,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Tenant Configuration -->
                <Border Grid.Row="0" Style="{StaticResource SectionCard}" Margin="0,0,0,12">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <TextBlock Grid.Row="0" Text="Tenant Configuration" Style="{StaticResource SectionHeader}"/>

                        <!-- Configuration Fields with Validation Indicators -->
                        <Grid Grid.Row="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <!-- Tenant ID with Validation -->
                            <StackPanel Grid.Column="0" Margin="0,0,12,0">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Tenant ID *" FontWeight="SemiBold" FontSize="13" Margin="0,0,0,6" VerticalAlignment="Center"/>
                                    <TextBlock Name="TenantValidation" Grid.Column="1" Style="{StaticResource ValidationIcon}"/>
                                </Grid>
                                <TextBox Name="TenantID" Style="{StaticResource ModernTextBox}"
                                         Text="company.onmicrosoft.com"
                                         Foreground="Gray"
                                         ToolTip="Your organization's Azure AD tenant identifier (e.g., contoso.onmicrosoft.com)"/>
                            </StackPanel>

                            <!-- Client ID with Validation -->
                            <StackPanel Grid.Column="1" Margin="12,0">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Client ID *" FontWeight="SemiBold" FontSize="13" Margin="0,0,0,6" VerticalAlignment="Center"/>
                                    <TextBlock Name="ClientValidation" Grid.Column="1" Style="{StaticResource ValidationIcon}"/>
                                </Grid>
                                <TextBox Name="ClientID" Style="{StaticResource ModernTextBox}"
                                         Text="14d82eec-204b-4c2f-b7e8-296a70dab67e"
                                         ToolTip="The Application (client) ID for your Azure AD app registration"/>
                            </StackPanel>

                            <!-- Config Status & Connect Button -->
                            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="12,0,0,0">
                                <TextBlock Name="ConfigStatus" VerticalAlignment="Center" Margin="0,0,16,0" 
                                     FontSize="12" Foreground="#6C757D" Text="Not connected"/>
                                <Button Name="ConnectGraphButton" Content="Connect" Style="{StaticResource PrimaryButton}" Width="120" Height="42"/>
                            </StackPanel>
                        </Grid>
                    </Grid>
                </Border>

                <!-- Enhanced Import List Panel with Tooltips -->
                <Border Grid.Row="1" Style="{StaticResource SectionCard}" Margin="0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <StackPanel Grid.Row="0" Orientation="Horizontal">
                            <TextBlock Text="Intune Import List" Style="{StaticResource SectionHeader}"/>
                            <TextBlock Name="ImportListCount" Text="(0 packages)" Foreground="#6C757D" 
                                     VerticalAlignment="Bottom" Margin="12,0,0,12" FontSize="12"/>
                            <TextBlock Name="ColumnHelp" Text="Column Help" 
                                       Foreground="#6C757D" TextDecorations="Underline" 
                                       Cursor="Hand" FontSize="11" VerticalAlignment="Bottom" Margin="12,0,0,13"/>
                        </StackPanel>

                        <!-- Main Action Buttons -->
                        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,12">
                            <Button Name="ImportToIntune" Content="Import to Intune" Style="{StaticResource SuccessButton}"/>
                            <Button Name="AdvancedSettings" Content="Advanced Settings" Style="{StaticResource SecondaryButton}"/>
                        </StackPanel>

                        <!-- Enhanced Import List DataGrid with Tooltips -->
                        <DataGrid Name="ImportList" Grid.Row="2" Style="{StaticResource ModernDataGrid}">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Package ID" Binding="{Binding PackageID}" Width="180">
                                    <DataGridTextColumn.ElementStyle>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="ToolTip" Value="The WinGet package identifier"/>
                                        </Style>
                                    </DataGridTextColumn.ElementStyle>
                                </DataGridTextColumn>
                                <DataGridComboBoxColumn Header="Context" SelectedValueBinding="{Binding Context}" Width="100">
                                    <DataGridComboBoxColumn.ItemsSource>
                                        <x:Array Type="sys:String">
                                            <sys:String>Machine</sys:String>
                                            <sys:String>User</sys:String>
                                            <sys:String>None</sys:String>
                                        </x:Array>
                                    </DataGridComboBoxColumn.ItemsSource>
                                    <DataGridComboBoxColumn.ElementStyle>
                                        <Style TargetType="ComboBox">
                                            <Setter Property="ToolTip" Value="Installation context: Machine (system-wide) or User (current user only)"/>
                                        </Style>
                                    </DataGridComboBoxColumn.ElementStyle>
                                </DataGridComboBoxColumn>
                                <DataGridComboBoxColumn Header="Accept Newer" SelectedValueBinding="{Binding AcceptNewerVersion}" Width="110">
                                    <DataGridComboBoxColumn.ItemsSource>
                                        <x:Array Type="sys:String">
                                            <sys:String>1</sys:String>
                                            <sys:String>0</sys:String>
                                        </x:Array>
                                    </DataGridComboBoxColumn.ItemsSource>
                                    <DataGridComboBoxColumn.ElementStyle>
                                        <Style TargetType="ComboBox">
                                            <Setter Property="ToolTip" Value="Accept newer versions during update (1=Yes, 0=No)"/>
                                        </Style>
                                    </DataGridComboBoxColumn.ElementStyle>
                                </DataGridComboBoxColumn>
                                <DataGridComboBoxColumn Header="Update Only" SelectedValueBinding="{Binding UpdateOnly}" Width="100">
                                    <DataGridComboBoxColumn.ItemsSource>
                                        <x:Array Type="sys:String">
                                            <sys:String>0</sys:String>
                                            <sys:String>1</sys:String>
                                        </x:Array>
                                    </DataGridComboBoxColumn.ItemsSource>
                                    <DataGridComboBoxColumn.ElementStyle>
                                        <Style TargetType="ComboBox">
                                            <Setter Property="ToolTip" Value="Only update existing installations (1=Yes, 0=No)"/>
                                        </Style>
                                    </DataGridComboBoxColumn.ElementStyle>
                                </DataGridComboBoxColumn>
                                <DataGridTextColumn Header="Target Version" Binding="{Binding TargetVersion}" Width="*">
                                    <DataGridTextColumn.ElementStyle>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="ToolTip" Value="Specific version to install (leave empty for latest)"/>
                                        </Style>
                                    </DataGridTextColumn.ElementStyle>
                                </DataGridTextColumn>
                                <DataGridTemplateColumn Header="Actions" Width="100">
                                    <DataGridTemplateColumn.CellTemplate>
                                        <DataTemplate>
                                            <Button Content="Advanced" 
                                                    Style="{StaticResource SecondaryButton}" 
                                                    Padding="8,4" FontSize="11" Height="28"
                                                    ToolTip="Edit advanced settings for this package"/>
                                        </DataTemplate>
                                    </DataGridTemplateColumn.CellTemplate>
                                </DataGridTemplateColumn>
                            </DataGrid.Columns>
                        </DataGrid>

                        <!-- Enhanced Action Buttons with Clear Visual Hierarchy -->
                        <Grid Grid.Row="3" Margin="0,16,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <!-- File Operations -->
                            <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left">
                                <Button Name="ImportCSV" Content="Import from CSV..." Style="{StaticResource SecondaryButton}"/>
                                <Button Name="ExportCSV" Content="Export to CSV..." Style="{StaticResource SecondaryButton}"/>
                            </StackPanel>

                            <!-- Critical Actions -->
                            <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button Name="DeleteSelected" Content="Delete Selected" Style="{StaticResource DangerButton}"/>
                                <Button Name="ClearAll" Content="Clear All" Style="{StaticResource DangerButton}"/>
                            </StackPanel>
                        </Grid>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <!-- SECTION 4: CONSOLE OUTPUT -->
        <Border Grid.Row="1" Style="{StaticResource SectionCard}">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <StackPanel Grid.Row="0" Orientation="Horizontal">
                    <TextBlock Text="Console Output" Style="{StaticResource SectionHeader}"/>
                    <Button Name="ClearConsole" Content="Clear" Style="{StaticResource SecondaryButton}" 
                          Width="80" Height="32" FontSize="11" Margin="16,0,0,16"/>
                </StackPanel>
                
                <Border Grid.Row="1"
                        Background="#FAFBFC" 
                        BorderBrush="#E1E5E9" BorderThickness="1" CornerRadius="4">
                    <ScrollViewer Name="ConsoleScrollViewer" 
                                VerticalScrollBarVisibility="Auto" 
                                HorizontalScrollBarVisibility="Auto">
                        <TextBox Name="ConsoleOutput" IsReadOnly="True" 
                               Background="Transparent" BorderThickness="0"
                               FontFamily="Consolas" FontSize="11"
                               TextWrapping="Wrap" AcceptsReturn="True"
                               Padding="12" VerticalAlignment="Top"/>
                    </ScrollViewer>
                </Border>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Load XAML with enhanced error handling
try {
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Error "Failed to load XAML: $_"
    exit 1
}

# Get WPF controls with enhanced organization
$controls = @{
    # Search Controls
    SearchBox           = $window.FindName("SearchBox")
    SearchButton        = $window.FindName("SearchButton")
    SearchStatus        = $window.FindName("SearchStatus")
    QuickFilter         = $window.FindName("QuickFilter")
    SearchResults       = $window.FindName("SearchResults")
    
    # Selection Controls
    
    # Transfer Controls
    TransferButton      = $window.FindName("TransferButton")
    TransferCount       = $window.FindName("TransferCount")
    
    # Import List Controls
    ImportList          = $window.FindName("ImportList")
    ImportListCount     = $window.FindName("ImportListCount")
    
    # Package Action Controls
    GetPackageDetails   = $window.FindName("GetPackageDetails")
    GetPackageVersions  = $window.FindName("GetPackageVersions")
    ViewOnGitHub        = $window.FindName("ViewOnGitHub")
    
    # File Operation Controls
    ImportCSV           = $window.FindName("ImportCSV")
    ExportCSV           = $window.FindName("ExportCSV")
    DeleteSelected      = $window.FindName("DeleteSelected")
    ClearAll            = $window.FindName("ClearAll")
    AdvancedSettings    = $window.FindName("AdvancedSettings")
    
    # Configuration Controls
    TenantID            = $window.FindName("TenantID")
    ClientID            = $window.FindName("ClientID")
    TenantValidation    = $window.FindName("TenantValidation")
    ClientValidation    = $window.FindName("ClientValidation")
    ConfigStatus        = $window.FindName("ConfigStatus")
    ConnectGraphButton  = $window.FindName("ConnectGraphButton")
    ImportToIntune      = $window.FindName("ImportToIntune")
    
    # Console Controls
    ConsoleOutput       = $window.FindName("ConsoleOutput")
    ConsoleScrollViewer = $window.FindName("ConsoleScrollViewer")
    ClearConsole        = $window.FindName("ClearConsole")
    
    # Help Controls
    GitHubLink          = $window.FindName("GitHubLink")
    ColumnHelp          = $window.FindName("ColumnHelp")
    SelectAllCheckBox   = $window.FindName("SelectAllCheckBox")
}

# Enhanced Console Output Function
function Write-ConsoleTextBox {
    param (
        [string]$Message,
        [switch]$NoTimeStamp,
        [string]$Type = "Info" # Info, Success, Warning, Error
    )

    if (-not $NoTimeStamp) {
        $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $Message = "[$TimeStamp] $Message"
    }

    # Add type indicators (fixed emoji encoding)
    switch ($Type) {
        "Success" { $Message = "SUCCESS: $Message" }
        "Warning" { $Message = "WARNING: $Message" }
        "Error" { $Message = "ERROR: $Message" }
        default { $Message = "INFO: $Message" }
    }

    # Output to console
    Write-Host $Message
    
    # Append to WPF console with auto-scroll
    $window.Dispatcher.Invoke([System.Action] {
            $controls.ConsoleOutput.AppendText("$Message`r`n")
            $controls.ConsoleScrollViewer.ScrollToBottom()
        })
}

# Initialize script-level variables
$script:originalSearchResults = @()
$script:importListItems = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]

# --- Asynchronous Job Handling ---
# The job-based asynchronous search is being replaced with a simpler synchronous model to ensure reliability.
# The related timer and job results hashtable are no longer needed for the search functionality.
# $script:jobResults = [System.Collections.Hashtable]::Synchronized(@{})
# $jobCheckTimer = New-Object System.Windows.Threading.DispatcherTimer
# $jobCheckTimer.Interval = [TimeSpan]::FromMilliseconds(200)
# $jobCheckTimer.Add_Tick({
#     # Check for completed search jobs
#     if ($script:jobResults.ContainsKey("Search")) {
#         $packages = $script:jobResults.Search
#         $script:jobResults.Remove("Search")

#         # Update UI on the main thread
#         $script:originalSearchResults = $packages
#         $controls.SearchResults.ItemsSource = $packages
#         $controls.SearchStatus.Text = "Found $($packages.Count) packages"
#         $controls.SearchStatus.Foreground = "#28A745"
#         $controls.SearchButton.IsEnabled = $true
#         Write-consoleTextBox "Found $($packages.Count) packages." -Type "Success"
#     }
#     # You can add checks for other job types here (e.g., Intune Import)
# })
# $jobCheckTimer.Start()
# --- End Asynchronous Job Handling ---


# Search Button Event Handler
$controls.SearchButton.Add_Click({ PerformSearch })

# Enhanced Search Function with Visual Feedback and Synchronous Execution
function PerformSearch {
    $searchTerm = $controls.SearchBox.Text.Trim()
    
    # Validate search term
    if ([string]::IsNullOrWhiteSpace($searchTerm) -or $searchTerm -eq $defaultSearchText) {
        Write-ConsoleTextBox "Please enter a valid search term." -Type "Warning"
        $controls.SearchStatus.Text = "Please enter a search term"
        $controls.SearchStatus.Foreground = "#DC3545"
        return
    }
    
    Write-ConsoleTextBox "--- Starting New Search ---"
    Write-ConsoleTextBox "Search term: '$searchTerm'" -Type "Info"
    
    # Update UI to indicate searching is in progress
    $controls.SearchStatus.Text = "Searching... Please wait."
    $controls.SearchStatus.Foreground = "#0078D7"
    $controls.SearchButton.IsEnabled = $false
    
    # Force UI to update before starting the synchronous search
    $window.Dispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

    try {
        Write-ConsoleTextBox "Searching for '$searchTerm' using the Microsoft.WinGet.Client module..." -Type "Info"
        
        # Use the PowerShell module for a more reliable search
        $searchResults = Find-WinGetPackage -Name $searchTerm -Source winget
        
        $packages = @()
        if ($searchResults) {
            foreach ($result in $searchResults) {
                $packageObject = [PSCustomObject]@{
                    Name       = $result.Name
                    ID         = $result.Id
                    Version    = $result.Version
                    IsSelected = $false
                }
                $packages += $packageObject
            }
        }
        
        Write-ConsoleTextBox "Successfully processed $($packages.Count) packages." -Type "Success"

        # Update UI on the main thread
        $script:originalSearchResults = $packages
        $controls.SearchResults.ItemsSource = $null # Clear old results first
        $controls.SearchResults.ItemsSource = $packages
        $controls.SearchStatus.Text = "Found $($packages.Count) packages"
        $controls.SearchStatus.Foreground = "#28A745"
    }
    catch {
        $errorMessage = "Error during search: $($_.Exception.Message)"
        Write-ConsoleTextBox $errorMessage -Type "Error"
        $controls.SearchStatus.Text = "Search error!"
        $controls.SearchStatus.Foreground = "#DC3545"
    }
    finally {
        # Re-enable the search button regardless of outcome
        $controls.SearchButton.IsEnabled = $true
        Write-ConsoleTextBox "--- Search Finished ---"
    }
}

# Enhanced Quick Filter Function
function ApplyQuickFilter {
    $filterText = $controls.QuickFilter.Text.Trim()
    
    # Store original results if not already stored
    if (-not $script:originalSearchResults -and $controls.SearchResults.ItemsSource) {
        $script:originalSearchResults = @($controls.SearchResults.ItemsSource)
    }
    
    if ([string]::IsNullOrWhiteSpace($filterText) -or $filterText -eq $defaultQuickFilterText) {
        # Show all items
        if ($script:originalSearchResults) {
            $controls.SearchResults.ItemsSource = $script:originalSearchResults
        }
    }
    else {
        # Filter items
        if ($script:originalSearchResults) {
            $filteredItems = $script:originalSearchResults | Where-Object {
                $_.Name -match $filterText -or $_.ID -match $filterText
            }
            $controls.SearchResults.ItemsSource = $filteredItems
        }
    }
    
    # Update transfer count after filtering
    UpdateTransferCount
}

# Enhanced Selection Management
function UpdateTransferCount {
    if ($controls.SearchResults.ItemsSource) {
        $selectedCount = @($controls.SearchResults.ItemsSource | Where-Object { $_.IsSelected }).Count
        $controls.TransferCount.Text = if ($selectedCount -gt 0) { "($selectedCount selected)" } else { "" }
    }
}

function SelectAllPackages {
    if ($controls.SearchResults.ItemsSource) {
        foreach ($item in $controls.SearchResults.ItemsSource) {
            $item.IsSelected = $true
        }
        $controls.SearchResults.Items.Refresh()
        UpdateTransferCount
        Write-ConsoleTextBox "Selected all visible packages" -Type "Success"
    }
}

function DeselectAllPackages {
    if ($controls.SearchResults.ItemsSource) {
        foreach ($item in $controls.SearchResults.ItemsSource) {
            $item.IsSelected = $false
        }
        $controls.SearchResults.Items.Refresh()
        UpdateTransferCount
        Write-ConsoleTextBox "Deselected all packages" -Type "Success"
    }
}

# Enhanced Import List Management
function UpdateImportListCount {
    $count = $script:importListItems.Count
    $controls.ImportListCount.Text = "($count packages)"
    $controls.ImportListCount.Foreground = if ($count -gt 0) { "#28A745" } else { "#6C757D" }
}

# Enhanced Configuration Validation
function ValidateConfiguration {
    $isValid = $true
    
    # Validate Tenant ID
    $tenantText = $controls.TenantID.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($tenantText) -or $tenantText -eq $defaultTenantText -or -not ($tenantText -like "*.*")) {
        $controls.TenantValidation.Text = "X"
        $controls.TenantValidation.Foreground = "#DC3545"
        $isValid = $false
    }
    else {
        $controls.TenantValidation.Text = ""
    }
    
    # Validate Client ID (basic GUID format check)
    $clientText = $controls.ClientID.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($clientText) -or $clientText.Length -lt 36) {
        $controls.ClientValidation.Text = "X"
        $controls.ClientValidation.Foreground = "#DC3545"
        $isValid = $false
    }
    else {
        $controls.ClientValidation.Text = ""
    }
    
    # Update status and button
    if ($isValid) {
        # Enable Import button only if connected and list has items
        $isConnected = try { $null -ne (Get-MSIntuneGraphAccessToken -ErrorAction SilentlyContinue) } catch { $false }
        if (-not $isConnected) {
            $controls.ConfigStatus.Text = "Not connected"
            $controls.ConfigStatus.Foreground = "#6C757D"
        }
        $controls.ImportToIntune.IsEnabled = ($script:importListItems.Count -gt 0) -and $isConnected
    }
    else {
        $controls.ConfigStatus.Text = "Invalid configuration"
        $controls.ConfigStatus.Foreground = "#DC3545"
        $controls.ImportToIntune.IsEnabled = $false
    }
}

# Function to connect to or disconnect from Microsoft Graph
function Toggle-GraphConnection {
    # If the button says "Disconnect", we disconnect
    if ($controls.ConnectGraphButton.Content -eq "Disconnect") {
        try {
            Disconnect-MSIntuneGraph
            Write-ConsoleTextBox "Successfully disconnected from Microsoft Graph." -Type "Success"
            $controls.ConfigStatus.Text = "Not connected"
            $controls.ConfigStatus.Foreground = "#6C757D"
            $controls.ConnectGraphButton.Content = "Connect"
            $controls.ConnectGraphButton.Style = $window.FindResource("PrimaryButton")
            # Re-enable text boxes
            $controls.TenantID.IsEnabled = $true
            $controls.ClientID.IsEnabled = $true
            $controls.ImportToIntune.IsEnabled = $false
        }
        catch {
            Write-ConsoleTextBox "Error during disconnect: $($_.Exception.Message)" -Type "Error"
        }
        return
    }

    # Otherwise, we connect
    ValidateConfiguration
    if ($controls.ConfigStatus.Text -eq "Invalid configuration") {
        Write-ConsoleTextBox "Cannot connect. Please fix configuration errors first." -Type "Warning"
        return
    }

    try {
        Write-ConsoleTextBox "Attempting to connect to Microsoft Graph via IntuneWin32App module..." -Type "Info"
        $redirectUrl = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        Connect-MSIntuneGraph -TenantID $controls.TenantID.Text -ClientID $controls.ClientID.Text -RedirectUri $redirectUrl -Interactive
        
        # Check if connected by trying to get context
        $context = Get-MSIntuneGraphAccessToken
        if ($context) {
            Write-ConsoleTextBox "Successfully connected to tenant '$($controls.TenantID.Text)'" -Type "Success"
            $controls.ConfigStatus.Text = "Connected to $($controls.TenantID.Text)"
            $controls.ConfigStatus.Foreground = "#28A745"
            $controls.ConnectGraphButton.Content = "Disconnect"
            $controls.ConnectGraphButton.Style = $window.FindResource("DangerButton")
            # Disable text boxes to prevent changes while connected
            $controls.TenantID.IsEnabled = $false
            $controls.ClientID.IsEnabled = $false
            # Re-validate to enable the import button
            ValidateConfiguration
        }
        else {
            throw "Connection verification failed."
        }
    }
    catch {
        Write-ConsoleTextBox "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Type "Error"
        $controls.ConfigStatus.Text = "Connection Failed"
        $controls.ConfigStatus.Foreground = "#DC3545"
    }
}

# Function to get WinGet package details
function WinGetPackageDetails {
    param (
        [string]$PackageID
    )

    try {
        Write-ConsoleTextBox "Retrieving package details for: $PackageID" -Type "Info"
        # Use winget.exe to get details from the repository
        $packageDetails = winget show --id $PackageID --source WinGet --accept-source-agreements --disable-interactivity --locale en-US
        
        if (-not $packageDetails) {
            Write-ConsoleTextBox "Package with ID '$PackageID' not found." -Type "Warning"
            return
        }

        # Format the output into a readable string for the popup and console
        $detailsString = $packageDetails -join "`r`n"
        $detailsString = $detailsString.Trim()

        Write-ConsoleTextBox "$PackageID - Details:" -Type "Info"
        $detailsString -split "`r?`n" | ForEach-Object { 
            Write-ConsoleTextBox $_ -NoTimeStamp 
        }
        Write-ConsoleTextBox "Details completed for $PackageID" -Type "Success"

        Show-DetailsPopup -Title "Package Details: $PackageID" -Content $detailsString

        return $detailsString
    }
    catch {
        Write-ConsoleTextBox "Error getting package details for $PackageID`: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to get WinGet package versions
function WinGetPackageVersions {
    param (
        [string]$PackageID
    )
    
    try {
        Write-ConsoleTextBox "Retrieving available versions for: $PackageID" -Type "Info"
        # Use winget.exe to get version info from the repository
        $packageVersionsOutput = winget show --id $PackageID --source WinGet --versions --accept-source-agreements --disable-interactivity --locale en-US

        if (-not $packageVersionsOutput) {
            Write-ConsoleTextBox "Package with ID '$PackageID' not found or has no other versions." -Type "Warning"
            return
        }

        Write-ConsoleTextBox "$PackageID - Available Versions:" -Type "Info"

        # Initialize an array for the available versions
        $WinGetPackageVersions = @()
        $versionLines = $packageVersionsOutput -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        # Skip header lines and process the remaining lines
        foreach ($line in $versionLines[2..($versionLines.Length - 1)]) {
            $trimmedLine = $line.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmedLine) -and $trimmedLine -notmatch "^-+$") {
                $WinGetPackageVersions += $trimmedLine
                Write-ConsoleTextBox $trimmedLine -NoTimeStamp
            }
        }

        Show-DetailsPopup -Title "Available Versions: $PackageID" -Content ($packageVersionsOutput -join "`r`n")

        return $WinGetPackageVersions
    }
    catch {
        Write-ConsoleTextBox "Error getting package versions for $PackageID`: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to show advanced settings popup
function Show-AdvancedSettingsPopup {
    param(
        [PSObject]$PackageItem
    )

    $popupXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"
        Title="Advanced Settings - $($PackageItem.PackageID)" 
        Height="600" Width="900"
        Background="#F8F9FA" 
        FontFamily="Segoe UI" 
        FontSize="12"
        WindowStartupLocation="CenterOwner"
        ResizeMode="CanResize"
        ShowInTaskbar="False">
    
    <Window.Resources>
        <!-- Button Styles -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#28A745"/>
        </Style>
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#6C757D"/>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="BorderBrush" Value="#DDE2E8"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Height" Value="42"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Height="{TemplateBinding Height}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="12,0,12,0" VerticalAlignment="Center" VerticalScrollBarVisibility="Hidden" HorizontalScrollBarVisibility="Hidden"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Card Style -->
        <Style x:Key="SectionCard" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E1E5E9"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
            <Setter Property="Margin" Value="8"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Margin="0,0,0,20">
            <StackPanel>
                <!-- Process Management -->
                <TextBlock Text="Process Management" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <Border Style="{StaticResource SectionCard}" Padding="15" Margin="0,0,0,15">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <TextBlock Text="Stop Process (Install)" Margin="0,0,0,5"/>
                            <TextBox Name="StopProcessInstall" Text="{Binding StopProcessInstall, Mode=TwoWay}" Style="{StaticResource ModernTextBox}"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="10,0,0,0">
                            <TextBlock Text="Stop Process (Uninstall)" Margin="0,0,0,5"/>
                            <TextBox Name="StopProcessUninstall" Text="{Binding StopProcessUninstall, Mode=TwoWay}" Style="{StaticResource ModernTextBox}"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Custom Arguments -->
                <TextBlock Text="Custom Arguments" FontSize="14" FontWeight="SemiBold" Margin="0,10,0,10"/>
                <Border Style="{StaticResource SectionCard}" Padding="15" Margin="0,0,0,15">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <TextBlock Text="Custom Args (Install)" Margin="0,0,0,5"/>
                            <TextBox Name="CustomArgumentListInstall" Text="{Binding CustomArgumentListInstall, Mode=TwoWay}" Style="{StaticResource ModernTextBox}"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="10,0,0,0">
                            <TextBlock Text="Custom Args (Uninstall)" Margin="0,0,0,5"/>
                            <TextBox Name="CustomArgumentListUninstall" Text="{Binding CustomArgumentListUninstall, Mode=TwoWay}" Style="{StaticResource ModernTextBox}"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Other Settings -->
                <TextBlock Text="Other Settings" FontSize="14" FontWeight="SemiBold" Margin="0,10,0,10"/>
                <Border Style="{StaticResource SectionCard}" Padding="15" Margin="0,0,0,15">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <TextBlock Text="Install Intent" Margin="0,0,0,5"/>
                            <TextBox Name="InstallIntent" Text="{Binding InstallIntent, Mode=TwoWay}" Style="{StaticResource ModernTextBox}"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="10,0,10,0">
                            <TextBlock Text="Group ID" Margin="0,0,0,5"/>
                            <TextBox Name="GroupID" Text="{Binding GroupID, Mode=TwoWay}" Style="{StaticResource ModernTextBox}"/>
                        </StackPanel>
                        <StackPanel Grid.Column="2" Margin="10,0,0,0">
                            <TextBlock Text="Show Notification" Margin="0,0,0,5"/>
                            <ComboBox Name="Notification" SelectedValue="{Binding Notification, Mode=TwoWay}" Height="38" VerticalContentAlignment="Center">
                                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">showAll</sys:String>
                                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">showReboot</sys:String>
                                <sys:String xmlns:sys="clr-namespace:System;assembly=mscorlib">hideAll</sys:String>
                            </ComboBox>
                        </StackPanel>
                    </Grid>
                </Border>
            </StackPanel>
        </ScrollViewer>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="SaveButton" Content="Save" Width="120" Style="{StaticResource SuccessButton}"/>
            <Button Name="CancelButton" Content="Cancel" Width="120" Style="{StaticResource SecondaryButton}"/>
        </StackPanel>
    </Grid>
</Window>
"@
    try {
        $popupReader = New-Object System.Xml.XmlNodeReader ([xml]$popupXaml)
        $popupWindow = [Windows.Markup.XamlReader]::Load($popupReader)
        $popupWindow.Owner = $window
        $popupWindow.DataContext = $PackageItem

        $saveButton = $popupWindow.FindName("SaveButton")
        $cancelButton = $popupWindow.FindName("CancelButton")

        $saveButton.Add_Click({
            $popupWindow.DialogResult = $true
            $popupWindow.Close()
        })

        $cancelButton.Add_Click({
            $popupWindow.DialogResult = $false
            $popupWindow.Close()
        })

        if ($popupWindow.ShowDialog()) {
            Write-ConsoleTextBox "Advanced settings updated for $($PackageItem.PackageID)." -Type "Success"
            return $true
        } else {
            Write-ConsoleTextBox "Advanced settings changes cancelled for $($PackageItem.PackageID)." -Type "Info"
            return $false
        }
    }
    catch {
        Write-ConsoleTextBox "Error showing advanced settings popup: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Function to show details popup
function Show-DetailsPopup {
    param (
        [string]$Title,
        [string]$Content
    )

    $popupXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" 
        Height="600" Width="800"
        Background="#F8F9FA" 
        FontFamily="Segoe UI" 
        FontSize="12"
        WindowStartupLocation="CenterOwner"
        ResizeMode="CanResize"
        ShowInTaskbar="False">
    
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="White" BorderBrush="#E1E5E9" BorderThickness="1" 
                CornerRadius="4" Padding="12" Margin="0,0,0,12">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                <TextBox Name="PopupContent" IsReadOnly="True" 
                       Background="Transparent" BorderThickness="0"
                       FontFamily="Consolas" FontSize="11"
                       TextWrapping="NoWrap" AcceptsReturn="True"/>
            </ScrollViewer>
        </Border>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="CopyButton" Content="Copy" Width="100" Height="36"
                  Background="#0078D7" Foreground="White" BorderThickness="0"
                  Margin="0,0,8,0" Cursor="Hand"/>
            <Button Name="CloseButton" Content="Close" Width="100" Height="36"
                  Background="#6C757D" Foreground="White" BorderThickness="0"
                  Cursor="Hand"/>
        </StackPanel>
    </Grid>
</Window>
"@

    try {
        $popupReader = New-Object System.Xml.XmlNodeReader ([xml]$popupXaml)
        $popupWindow = [Windows.Markup.XamlReader]::Load($popupReader)
        
        $popupWindow.Owner = $window
        
        $popupContent = $popupWindow.FindName("PopupContent")
        $popupContent.Text = $Content
        $copyButton = $popupWindow.FindName("CopyButton")
        $closeButton = $popupWindow.FindName("CloseButton")
        
        $copyButton.Add_Click({
            [System.Windows.Clipboard]::SetText($popupContent.Text)
            Write-ConsoleTextBox "Content copied to clipboard" -Type "Success"
        })
        
        $closeButton.Add_Click({
            $popupWindow.Close()
        })
        
        $popupWindow.Add_KeyDown({
            param($sender, $e)
            if ($e.Key -eq "Escape") {
                $popupWindow.Close()
            }
        })
        
        $popupWindow.ShowDialog() | Out-Null
    }
    catch {
        Write-ConsoleTextBox "Error showing popup: $($_.Exception.Message)" -Type "Error"
    }
}

# Function to show column definitions help
function GetColumnDefinitions {
    $columnDefinitions = @"
Column Definitions for Intune Import List:

PackageID: The WinGet package identifier (e.g., Microsoft.VisualStudioCode)
Context: Installation context - 'Machine' for system-wide, 'User' for current user
Accept Newer: Accept newer versions during update (1 = Yes, 0 = No)
Update Only: Only update existing installations (1 = Yes, 0 = No)
Target Version: Specific version to install (leave empty for latest)
Stop Process (Install): Process names to stop before installation (comma-separated)
Stop Process (Uninstall): Process names to stop before uninstallation (comma-separated)
Pre Script Install: PowerShell script to run before installation
Post Script Install: PowerShell script to run after installation
Pre Script Uninstall: PowerShell script to run before uninstallation
Post Script Uninstall: PowerShell script to run after uninstallation
Custom Args (Install): Custom arguments for installation command
Custom Args (Uninstall): Custom arguments for uninstallation command
Install Intent: Installation intent for Intune deployment
Notification: Show notifications during installation (1 = Yes, 0 = No)
GroupID: Group identifier for batch operations
"@

    Show-DetailsPopup -Title "Column Definitions Help" -Content $columnDefinitions
}

# Function to read log file and update the GUI
function Update-GUIFromLogFile {
    param (
        [string]$logFilePath
    )
    
    try {
        if (Test-Path $logFilePath) {
            $logContent = Get-Content -Path $logFilePath
            Write-ConsoleTextBox "Loading log file content..." -Type "Info"
            
            foreach ($line in $logContent) {
                Write-ConsoleTextBox $line -NoTimeStamp
            }
            
            Write-ConsoleTextBox "Log file content loaded successfully" -Type "Success"
        }
        else {
            Write-ConsoleTextBox "Log file not found: $logFilePath" -Type "Warning"
        }
    }
    catch {
        Write-ConsoleTextBox "Error reading log file: $($_.Exception.Message)" -Type "Error"
    }
}

# Enhanced transfer functionality
$controls.TransferButton.Add_Click({
    if (-not $controls.SearchResults.ItemsSource) {
        Write-ConsoleTextBox "No search results available." -Type "Warning"
        return
    }

    $selectedItems = @($controls.SearchResults.ItemsSource | Where-Object { $_.IsSelected })

    if ($selectedItems.Count -eq 0) {
        Write-ConsoleTextBox "No packages selected. Please select packages to transfer." -Type "Warning"
        return
    }

    # Initialize import list if needed
    if (-not $controls.ImportList.ItemsSource) {
        $controls.ImportList.ItemsSource = $script:importListItems
    }

    $addedCount = 0
    foreach ($item in $selectedItems) {
        # Check for duplicates
        $existingItem = $script:importListItems | Where-Object { $_.PackageID -eq $item.ID }
        if ($existingItem) {
            Write-ConsoleTextBox "Package '$($item.ID)' already in import list" -Type "Warning"
            continue
        }
    
        # Create new import item with enhanced defaults
        $newItem = [PSCustomObject]@{
            PackageID                   = $item.ID
            Context                     = "Machine"
            AcceptNewerVersion          = "1"
            UpdateOnly                  = "0"
            TargetVersion               = ""
            StopProcessInstall          = ""
            StopProcessUninstall        = ""
            PreScriptInstall            = ""
            PostScriptInstall           = ""
            PreScriptUninstall          = ""
            PostScriptUninstall         = ""
            CustomArgumentListInstall   = ""
            CustomArgumentListUninstall = ""
            InstallIntent               = ""
            Notification                = "showAll"
            GroupID                     = ""
        }
    
        $script:importListItems.Add($newItem)
        $addedCount++
    
        # Deselect transferred item
        $item.IsSelected = $false
    }

    if ($addedCount -gt 0) {
        Write-ConsoleTextBox "Added $addedCount packages to import list" -Type "Success"
        $controls.SearchResults.Items.Refresh()
        UpdateTransferCount
        UpdateImportListCount
        ValidateConfiguration
    }
})

# Import list management - FIXED function calls
$controls.DeleteSelected.Add_Click({
    $selectedItems = @($controls.ImportList.SelectedItems)

    if ($selectedItems.Count -eq 0) {
        Write-ConsoleTextBox "No items selected for deletion." -Type "Warning"
        return
    }

    foreach ($item in $selectedItems) {
        $script:importListItems.Remove($item)
    }

    Write-ConsoleTextBox "Removed $($selectedItems.Count) items from import list" -Type "Success"
    UpdateImportListCount
    ValidateConfiguration
})

$controls.ClearAll.Add_Click({
    if ($script:importListItems.Count -eq 0) {
        Write-ConsoleTextBox "Import list is already empty." -Type "Warning"
        return
    }

    $count = $script:importListItems.Count
    $script:importListItems.Clear()
    Write-ConsoleTextBox "Cleared all $count items from import list" -Type "Success"
    UpdateImportListCount
    ValidateConfiguration
})

# Use a routed event handler for buttons inside the DataGrid
$controls.ImportList.AddHandler([System.Windows.Controls.Button]::ClickEvent, [System.Windows.RoutedEventHandler]{
    param($sender, $e)
    
    # Check if the clicked element is a Button
    if ($e.OriginalSource -is [System.Windows.Controls.Button]) {
        $button = $e.OriginalSource
        # Check if the button is our advanced settings button
        if ($button.Content -eq "Advanced") {
            # The DataContext of the button is the row's item
            $itemToEdit = $button.DataContext
            if ($itemToEdit) {
                # Create a clone for editing, so changes can be cancelled
                $itemClone = $itemToEdit.PSObject.Copy()
                if (Show-AdvancedSettingsPopup -PackageItem $itemClone) {
                    # If saved, update the original item's properties
                    foreach ($prop in $itemClone.psobject.Properties) {
                        if ($itemToEdit.psobject.Properties[$prop.Name]) {
                            $itemToEdit.psobject.Properties[$prop.Name].Value = $prop.Value
                        }
                    }
                    $controls.ImportList.Items.Refresh()
                    Write-ConsoleTextBox "Advanced settings saved for $($itemToEdit.PackageID)." -Type "Success"
                }
            }
        }
    }
})

# Function to open advanced settings for a package
function Open-AdvancedSettings {
    $selectedItem = $controls.ImportList.SelectedItem
    if ($selectedItem) {
        # Create a clone for editing, so changes can be cancelled
        $itemToEdit = $selectedItem.PSObject.Copy()
        if (Show-AdvancedSettingsPopup -PackageItem $itemToEdit) {
            # If saved, update the original item's properties by iterating through the clone's properties
            foreach ($prop in $itemToEdit.psobject.Properties) {
                if ($selectedItem.psobject.Properties[$prop.Name]) {
                    $selectedItem.psobject.Properties[$prop.Name].Value = $prop.Value
                }
            }
            $controls.ImportList.Items.Refresh()
            Write-ConsoleTextBox "Advanced settings saved for $($selectedItem.PackageID)." -Type "Success"
        }
    } else {
        Write-ConsoleTextBox "Please select a package from the import list to edit." -Type "Warning"
    }
}

# Function to open the package manifest page on GitHub
function Show-PackageOnGitHub {
    $selectedItem = $controls.SearchResults.SelectedItem
    if (-not $selectedItem) {
        Write-ConsoleTextBox "Please select a package to view on GitHub." -Type "Warning"
        return
    }

    try {
        $packageId = $selectedItem.ID
        $firstChar = $packageId.Substring(0, 1).ToLower()
        
        $firstDotIndex = $packageId.IndexOf('.')
        if ($firstDotIndex -le 0) {
            throw "Package ID '$packageId' is not in the expected 'Publisher.AppName' format."
        }
        
        $publisher = $packageId.Substring(0, $firstDotIndex)
        $appName = $packageId.Substring($firstDotIndex + 1)

        $gitHubUrl = "https://github.com/microsoft/winget-pkgs/tree/master/manifests/$firstChar/$publisher/$appName"
        
        Write-ConsoleTextBox "Opening GitHub URL: $gitHubUrl" -Type "Info"
        Start-Process $gitHubUrl
    }
    catch {
        Write-ConsoleTextBox "Could not construct GitHub URL for '$($selectedItem.ID)': $($_.Exception.Message)" -Type "Error"
    }
}

# Event handler for the main "Advanced Settings" button
$controls.AdvancedSettings.Add_Click({
    Open-AdvancedSettings
})

# Event handler for selection changes in the import list
$controls.ImportList.Add_SelectionChanged({
    # Enable/disable buttons based on selection
    $hasSelection = $controls.ImportList.SelectedItems.Count -gt 0
    $controls.AdvancedSettings.IsEnabled = $hasSelection
    $controls.DeleteSelected.IsEnabled = $hasSelection
})

# Package details buttons
$controls.GetPackageDetails.Add_Click({
    $selectedItem = $controls.SearchResults.SelectedItem
    if ($selectedItem) {
        WinGetPackageDetails -PackageID $selectedItem.ID
    } else {
        Write-ConsoleTextBox "Please select a package to view details." -Type "Warning"
    }
})

$controls.GetPackageVersions.Add_Click({
    $selectedItem = $controls.SearchResults.SelectedItem
    if ($selectedItem) {
        WinGetPackageVersions -PackageID $selectedItem.ID
    } else {
        Write-ConsoleTextBox "Please select a package to view versions." -Type "Warning"
    }
})

$controls.ViewOnGitHub.Add_Click({
    Show-PackageOnGitHub
})

# Quick filter functionality - FIXED function call
$controls.QuickFilter.Add_TextChanged({ ApplyQuickFilter })

# Configuration validation on text changes - FIXED function calls
$controls.TenantID.Add_TextChanged({ ValidateConfiguration })
$controls.ClientID.Add_TextChanged({ ValidateConfiguration })

# File operations with fixed function calls
$controls.ImportCSV.Add_Click({
    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openFileDialog.InitialDirectory = (Get-Location).Path
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Import Package List from CSV"

    if ($openFileDialog.ShowDialog()) {
        try {
            $importedData = Import-Csv -Path $openFileDialog.FileName
        
            $script:importListItems.Clear()
            foreach ($item in $importedData) {
                $script:importListItems.Add($item)
            }
        
            if (-not $controls.ImportList.ItemsSource) {
                $controls.ImportList.ItemsSource = $script:importListItems
            }
        
            Write-ConsoleTextBox "Imported CSV: $($openFileDialog.FileName)" -Type "Success"
            UpdateImportListCount
            ValidateConfiguration
        }
        catch {
            Write-ConsoleTextBox "Error importing CSV: $($_.Exception.Message)" -Type "Error"
        }
    }
})

$controls.ExportCSV.Add_Click({
    if ($script:importListItems.Count -eq 0) {
        Write-ConsoleTextBox "No data to export." -Type "Warning"
        return
    }

    $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveFileDialog.InitialDirectory = (Get-Location).Path
    $saveFileDialog.FileName = "WinGet-WrapperImportGUI-$timestamp.csv"
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Export Package List to CSV"

    if ($saveFileDialog.ShowDialog()) {
        try {
            $script:importListItems | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            Write-ConsoleTextBox "Exported: $($saveFileDialog.FileName)" -Type "Success"
        }
        catch {
            Write-ConsoleTextBox "Error exporting CSV: $($_.Exception.Message)" -Type "Error"
        }
    }
})

# Console management
$controls.ClearConsole.Add_Click({
    $controls.ConsoleOutput.Clear()
    Write-ConsoleTextBox "Console cleared" -Type "Success"
})

# Help links
$controls.GitHubLink.Add_MouseDown({
    Start-Process "https://github.com/SorenLundt/WinGet-Wrapper"
})

$controls.ColumnHelp.Add_MouseDown({
    GetColumnDefinitions
})

# Enhanced placeholder handling for search box
$defaultSearchText = "Search for software (e.g., VLC, Chrome)"
$controls.SearchBox.Add_GotFocus({
    if ($controls.SearchBox.Text -eq $defaultSearchText) {
        $controls.SearchBox.Text = ""
        $controls.SearchBox.Foreground = "Black"
    }
})

$controls.SearchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($controls.SearchBox.Text)) {
        $controls.SearchBox.Text = $defaultSearchText
        $controls.SearchBox.Foreground = "Gray"
    }
})

# Enhanced placeholder handling for quick filter
$defaultQuickFilterText = "Quick filter results..."
$controls.QuickFilter.Add_GotFocus({
    if ($controls.QuickFilter.Text -eq $defaultQuickFilterText) {
        $controls.QuickFilter.Text = ""
        $controls.QuickFilter.Foreground = "Black"
    }
})

$controls.QuickFilter.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($controls.QuickFilter.Text)) {
        $controls.QuickFilter.Text = $defaultQuickFilterText
        $controls.QuickFilter.Foreground = "Gray"
    }
})

# Enhanced placeholder handling for tenant ID
$defaultTenantText = "company.onmicrosoft.com"
$controls.TenantID.Add_GotFocus({
    if ($controls.TenantID.Text -eq $defaultTenantText) {
        $controls.TenantID.Text = ""
        $controls.TenantID.Foreground = "Black"
    }
})

$controls.TenantID.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($controls.TenantID.Text)) {
        $controls.TenantID.Text = $defaultTenantText
        $controls.TenantID.Foreground = "Gray"
    }
})

# Graph Connection Button
$controls.ConnectGraphButton.Add_Click({
    Toggle-GraphConnection
})

# Enhanced Import to Intune functionality
$controls.ImportToIntune.Add_Click({
    Write-ConsoleTextBox "Starting import to Intune..." -Type "Info"

    # Check for active IntuneWin32App connection
    $isConnected = try { $null -ne (Get-MSIntuneGraphAccessToken -ErrorAction SilentlyContinue) } catch { $false }
    if (-not $isConnected) {
        Write-ConsoleTextBox "Not connected to Microsoft Graph. Please use the 'Connect' button first." -Type "Error"
        return
    }

    # Validate configuration
    ValidateConfiguration
    if (-not $controls.ImportToIntune.IsEnabled) {
        Write-ConsoleTextBox "Please check configuration before importing." -Type "Error"
        return
    }

    # Check required files
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
            Write-ConsoleTextBox "Required file missing: '$file'" -Type "Error"
            $foundAllFiles = $false
        }
        else {
            Write-ConsoleTextBox "Found required file: '$file'" -Type "Success"
        }
    }

    if (-not $foundAllFiles) {
        Write-ConsoleTextBox "Cannot proceed - missing required files." -Type "Error"
        return
    }

    if ($script:importListItems.Count -eq 0) {
        Write-ConsoleTextBox "No packages to import." -Type "Warning"
        return
    }

    try {
        # Export temporary CSV
        $fileName = "TempExport-$timestamp.csv"
        $csvFilePath = Join-Path -Path $scriptRoot -ChildPath $fileName
        $script:importListItems | Export-Csv -Path $csvFilePath -NoTypeInformation
        
        Write-ConsoleTextBox "Created temporary export: $csvFilePath" -Type "Success"

        # Prepare import script execution
        $logFile = "$scriptRoot\Logs\WinGet_WrapperImportFromCSV_$timestamp.log"
        $importScriptPath = Join-Path -Path $scriptRoot -ChildPath "Winget-WrapperImportFromCSV.ps1"
        
        $redirectUrl = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        $arguments = "-csvFile `"$csvFilePath`" -TenantID $($controls.TenantID.Text) -ClientID $($controls.ClientID.Text) -RedirectURL `"$redirectUrl`" -ScriptRoot `"$scriptRoot`" -SkipConfirmation -SkipModuleCheck -SkipGraphConnection"

        Write-ConsoleTextBox "Executing import script with arguments: $arguments"
                
        # Use -PassThru to get the process object and wait for it to exit
        $process = Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File `"$importScriptPath`"", $arguments -NoNewWindow -PassThru
        $process.WaitForExit()

        Write-ConsoleTextBox "Import script process has finished. Reading log..." -Type "Info"

        # Update GUI from log file immediately after process exit
        Update-GUIFromLogFile -logFilePath $logFile

        # Remove temporary CSV
        if (Test-Path $csvFilePath) {
            Remove-Item $csvFilePath -Force
            Write-ConsoleTextBox "Temporary file $csvFilePath deleted successfully." -Type "Success"
        }

        Write-ConsoleTextBox "Import process completed. Check log file: $logFile" -Type "Success"
    }
    catch {
        Write-ConsoleTextBox "Error during import: $($_.Exception.Message)" -Type "Error"
    }
})

# Add event handler for the new "Select All" checkbox
$controls.SelectAllCheckBox.Add_Click({
    param($sender, $e)
    $isChecked = $sender.IsChecked
    if ($controls.SearchResults.ItemsSource) {
        foreach ($item in $controls.SearchResults.ItemsSource) {
            $item.IsSelected = $isChecked
        }
        # Refresh the view to show the updated checkboxes
        $controls.SearchResults.Items.Refresh()
        UpdateTransferCount
    }
})


# Initialize the application
Write-ConsoleTextBox "WinGet-Wrapper Import GUI v2.1 initialized successfully" -Type "Success"
Write-ConsoleTextBox "Ready to search and import packages to Intune" -Type "Info"

# Initial validation and setup
ValidateConfiguration
UpdateImportListCount
$controls.AdvancedSettings.IsEnabled = $false
$controls.DeleteSelected.IsEnabled = $false


# Set initial focus to search box
$window.Add_Loaded({
    $controls.SearchBox.Focus()
})

# Show the window
$window.ShowDialog() | Out-Null