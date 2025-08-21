# Soren Lundt - 22-02-2023
# URL: https://github.com/SorenLundt/WinGet-Wrapper
# License: https://raw.githubusercontent.com/SorenLundt/WinGet-Wrapper/main/LICENSE.txt
# Wrapper for running winget in system context. Finds the WinGet install path and runs desired command
#
# Version 2.0 - 17-07-2025 - Enhanced with 64-bit enforcement, granular parameters, uninstall function, and process-specific logging

# Enforce 64-bit execution
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    try {
        $argsString = ""
        foreach ($k in $MyInvocation.BoundParameters.keys) {
            switch ($MyInvocation.BoundParameters[$k].GetType().Name) {
                "SwitchParameter" { if ($MyInvocation.BoundParameters[$k].IsPresent) { $argsString += "-$k " } }
                "String" { $argsString += "-$k `"$($MyInvocation.BoundParameters[$k])`" " }
                "Int32" { $argsString += "-$k $($MyInvocation.BoundParameters[$k]) " }
                "Boolean" { $argsString += "-$k `$$($MyInvocation.BoundParameters[$k]) " }
            }
        }
        Start-Process -FilePath "$env:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" $argsString" -Wait -NoNewWindow
        exit $LASTEXITCODE
    }
    catch {
        Write-Error "Failed to re-launch in 64-bit PowerShell. Error: $($_.Exception.Message)"
        exit 1
    }
}

#Parameters - Granular approach
Param (
    [Parameter(Mandatory=$true)]
    [string]$PackageID,

    [Parameter(Mandatory=$false)]
    [ValidateSet("Machine", "User", "None", "")]
    [string]$Scope,

    [Parameter()]
    [string]$StopProcess,

    [Parameter()]
    [string]$PreScript,

    [Parameter()]
    [string]$PostScript,
    
    [Parameter()]
    [string]$CustomArgs,

    [Parameter()]
    [string]$TargetVersion,

    [Parameter()]
    [switch]$UpdateOnly,

    [Parameter()]
    [switch]$Uninstall
)

#Define common variables
$ScriptVersion = "2.0"
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logPath = "$env:ProgramData\WinGet-WrapperLogs\$PackageID"
$stdout = "$logPath\StdOut-$timestamp.txt"
$errout = "$logPath\ErrOut-$timestamp.txt"

# Create process-specific log folder
if (!(Test-Path -Path $logPath)) {
    try {
        New-Item -Path $logPath -Force -ItemType Directory | Out-Null
    }
    catch {
        Write-Host "Failed to create log directory: $($_.Exception.Message)"
        exit 1
    }
}

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "$logPath\WinGet_Wrapper_$($timestamp).log"
    )

    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] $Message"

    # Output to the console
    Write-Host $LogEntry

    # Append to the log file
    $LogEntry | Out-File -Append -FilePath $LogFile
}

#Write useful variables to log
Write-Log "                    WinGet-Wrapper v.$ScriptVersion"
Write-Log "https://github.com/SorenLundt/WinGet-Wrapper"
Write-Log "                GNU General Public License v3"
Write-Log "**************************************************"
Write-Log "Host: $env:ComputerName"
Write-Log "PackageID: $PackageID"
Write-Log "Scope: $Scope"
Write-Log "Uninstall Switch: $($Uninstall.IsPresent)"
Write-Log "UpdateOnly Switch: $($UpdateOnly.IsPresent)"
Write-Log "TargetVersion: $TargetVersion"
Write-Log "StopProcess: $StopProcess"
Write-Log "PreScript: $PreScript"
Write-Log "PostScript: $PostScript"
Write-Log "CustomArgs: $CustomArgs"
Write-Log "LogPath: $logPath"
Write-Log "**************************************************"

#PreScript execution
if (![string]::IsNullOrEmpty($PreScript)) {
    Write-Log "Running PreScript: $PreScript"
    if (Test-Path $PreScript -PathType Leaf) {
        try {
            Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $PreScript
            Write-Log "PreScript completed successfully"
        } catch {
            Write-Log "An error occurred while running the PreScript: $_"
            Write-Log "Exit 1"
            exit 1
        }
    } else {
        Write-Log "PreScript not found: $PreScript"
        Write-Log "Exit 1"
        exit 1
    }
}

#Stop process if specified
if (-not ($StopProcess -eq $null) -and $StopProcess -ne "") {
    $processes = $StopProcess -split ","
    foreach ($process in $processes) {
        $processName = $process.Trim()
        if ($processName -ne "") {
            Stop-Process -Name $processName -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Write-Log "Stopped process: $processName"
        }
    }
}

#Determine if running in system or user context
if ($env:USERNAME -like "*$env:COMPUTERNAME*") {
    Write-Log "Running in System Context"
    $Context = "Machine"
}
else {
    Write-Log "Running in User Context"
    $Context = "User"
}

# Find WinGet.exe Location
if ($Context -contains "Machine"){
    try {
        $resolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*__8wekyb3d8bbwe"
        if ($resolveWingetPath) {
            $wingetPath = $resolveWingetPath[-1].Path
            $wingetPath = $wingetPath + "\winget.exe"
            Write-Log "WinGet path: $wingetPath"
        }
        else {
            Write-Log "Failed to find WinGet path"
            exit 1
        }
    }
    catch {
        Write-Log "Failed to find WinGet path: $($_.Exception.Message)"
        exit 1
    }
}
else{
    # Running in user context. Set WingetPath
    $wingetPath = "winget.exe" 
    Write-Log "Running in user context, using system PATH for WinGet"
}

# Build ArgumentList based on parameters and action
$ArgumentList = ""
if ($Uninstall.IsPresent) {
    # Uninstall logic
    $ArgumentList = "uninstall --id `"$PackageID`" --exact --silent --accept-source-agreements"
    if ($Scope -and $Scope -ne "None") {
        $ArgumentList += " --scope $Scope"
    }
    if (![string]::IsNullOrEmpty($TargetVersion)) {
        $ArgumentList += " --version `"$TargetVersion`""
    }
    if (![string]::IsNullOrEmpty($CustomArgs)) {
        $ArgumentList += " $CustomArgs"
    }
    Write-Log "Action: Uninstall"
}
else {
    # Install logic
    if ($UpdateOnly.IsPresent) {
        $ArgumentList = "update --id `"$PackageID`" --exact --silent --accept-package-agreements --accept-source-agreements"
        Write-Log "Action: Update"
    }
    else {
        $ArgumentList = "install --id `"$PackageID`" --exact --silent --accept-package-agreements --accept-source-agreements"
        Write-Log "Action: Install"
    }
    
    if ($Scope -and $Scope -ne "None") {
        $ArgumentList += " --scope $Scope"
    }
    if (![string]::IsNullOrEmpty($TargetVersion)) {
        $ArgumentList += " --version `"$TargetVersion`""
    }
    if (![string]::IsNullOrEmpty($CustomArgs)) {
        $ArgumentList += " $CustomArgs"
    }
}

Write-Log "Executing: $wingetPath $ArgumentList"

# Execute WinGet command
try {
    $process = Start-Process "$WingetPath" -ArgumentList "$ArgumentList" -WindowStyle Hidden -PassThru -Wait -RedirectStandardOutput "$stdout" -RedirectStandardError "$errout"
    $exitCode = $process.ExitCode
    Write-Log "WinGet process exit code: $exitCode"
}
catch {
    Write-Log "Error executing WinGet: $($_.Exception.Message)"
    exit 1
}

# Log the contents of stdout and stderr
if (Test-Path $stdout) {
    $stdoutContent = Get-Content "$stdout" -ErrorAction SilentlyContinue
    if ($stdoutContent) {
        Write-Log "--- WinGet Standard Output ---"
        $stdoutContent | ForEach-Object {
            Write-Log "STDOUT: $_"
        }
    }
    Remove-Item -Path "$stdout" -Force -ErrorAction SilentlyContinue
}

if (Test-Path $errout) {
    $erroutContent = Get-Content "$errout" -ErrorAction SilentlyContinue
    if ($erroutContent) {
        Write-Log "--- WinGet Standard Error ---"
        $erroutContent | ForEach-Object {
            Write-Log "STDERR: $_"
        }
    }
    Remove-Item -Path "$errout" -Force -ErrorAction SilentlyContinue
}

#PostScript execution
if (![string]::IsNullOrEmpty($PostScript)) {
    Write-Log "Running PostScript: $PostScript"
    if (Test-Path $PostScript -PathType Leaf) {
        try {
            Powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File $PostScript
            Write-Log "PostScript completed successfully"
        } catch {
            Write-Log "An error occurred while running the PostScript: $_"
            Write-Log "Exit 1"
            exit 1
        }
    } else {
        Write-Log "PostScript not found: $PostScript"
        Write-Log "Exit 1"
        exit 1
    }
}

Write-Log "Script Finished with exit code: $exitCode"
exit $exitCode