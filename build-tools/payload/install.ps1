param(
    [string]$InstallDir = "$env:APPDATA\VMWV.Modern",
    [switch]$RemoveNodeJsRuntime,
    [switch]$RemoveLegacy,
    [switch]$SkipExtract,
    [switch]$LaunchAfterInstall
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appName = "Voicemeeter Windows Volume Modern"
$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$appName"
$legacyTaskNames = @("voicemeeter-windows-volume")
$newTaskNames = @("VMWV.Modern")

# Log to dev-logs folder next to the installer script (or fallback to TEMP)
$devLogsDir = Join-Path (Split-Path $scriptDir -Parent | Split-Path -Parent) "dev-logs"
if (-not (Test-Path $devLogsDir)) {
    $devLogsDir = $env:TEMP
}
$logFile = Join-Path $devLogsDir ("install-" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")

function Write-Log {
    param([string]$Message)
    $line = "[$(Get-Date -Format 'HH:mm:ss')] [Installer] $Message"
    Write-Host $line
    Add-Content -Path $logFile -Value $line
}

function Stop-RunningProcesses {
    foreach ($name in @("VMWV", "VMWV.Modern")) {
        try {
            Get-Process -Name $name -ErrorAction SilentlyContinue | Stop-Process -Force
        } catch {
        }
    }
}

function Remove-TaskIfExists {
    param([string]$TaskName)

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    } catch {
    }
}

function Invoke-LegacyUninstall {
    $paths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $paths) {
        $entries = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
        foreach ($entry in $entries) {
            if (-not $entry.DisplayName) {
                continue
            }

            if ($entry.DisplayName -like "Voicemeeter Windows Volume*" -and $entry.DisplayName -notlike "*Modern*") {
                Write-Log "Found legacy install: $($entry.DisplayName)"

                $uninstallString = $entry.UninstallString
                if (-not $uninstallString) {
                    continue
                }

                try {
                    if ($uninstallString -match "(?i)uninstall\.exe") {
                        $exePath = $null
                        if ($uninstallString.StartsWith('"')) {
                            $exePath = $uninstallString.Split('"')[1]
                        } else {
                            $exePath = $uninstallString.Split(' ')[0]
                        }

                        if ($exePath -and (Test-Path $exePath)) {
                            Write-Log "Running legacy uninstaller silently"
                            Start-Process -FilePath $exePath -ArgumentList "/S" -Wait
                        }
                    } else {
                        Write-Log "Running legacy uninstall string"
                        Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $uninstallString -Wait
                    }
                } catch {
                    Write-Log "Legacy uninstall returned an error but install will continue: $($_.Exception.Message)"
                }
            }
        }
    }
}

function Remove-NodeJsRuntimeIfRequested {
    if (-not $RemoveNodeJsRuntime) {
        return
    }

    Write-Log "RemoveNodeJsRuntime enabled: attempting Node.js uninstall"

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        try {
            Start-Process -FilePath "winget" -ArgumentList "uninstall", "--id", "OpenJS.NodeJS", "-e", "--accept-source-agreements" -Wait
            Start-Process -FilePath "winget" -ArgumentList "uninstall", "--id", "OpenJS.NodeJS.LTS", "-e", "--accept-source-agreements" -Wait
        } catch {
            Write-Log "winget Node.js uninstall encountered an error: $($_.Exception.Message)"
        }
    }

    $msiKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($key in $msiKeys) {
        $entries = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        foreach ($entry in $entries) {
            if ($entry.DisplayName -like "Node.js*") {
                $uninstallString = $entry.UninstallString
                if (-not $uninstallString) {
                    continue
                }

                try {
                    Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $uninstallString, "/quiet" -Wait
                } catch {
                    Write-Log "Node.js uninstall entry failed: $($_.Exception.Message)"
                }
            }
        }
    }
}

function Ensure-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$TargetPath,
        [string]$WorkingDirectory
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.Save()
}

Write-Log "Starting install"
Write-Log "Log file: $logFile"
Write-Log "Install directory: $InstallDir"

try {

Stop-RunningProcesses

if ($RemoveLegacy) {
    Write-Log "Legacy removal requested"
    Invoke-LegacyUninstall
} else {
    Write-Log "Skipping legacy removal (not selected)"
}

Remove-NodeJsRuntimeIfRequested

foreach ($taskName in $legacyTaskNames + $newTaskNames) {
    Remove-TaskIfExists -TaskName $taskName
}

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

$payloadZip = Join-Path $scriptDir "payload.zip"

Write-Log "Extracting application payload"
if ($SkipExtract) {
    Write-Log "Skipping extraction (already extracted by NSIS)"
} else {
    $payloadZip = Join-Path $scriptDir "payload.zip"
    if (-not (Test-Path $payloadZip)) {
        throw "Missing payload.zip next to installer script"
    }
    Expand-Archive -Path $payloadZip -DestinationPath $InstallDir -Force
}

Write-Log "Payload ready"

$exePath = Join-Path $InstallDir "VMWV.Modern.exe"
if (-not (Test-Path $exePath)) {
    throw "Install completed but executable was not found at $exePath"
}

$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\VMWV.Modern"
New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null
Ensure-Shortcut -ShortcutPath (Join-Path $startMenuDir "VMWV.Modern.lnk") -TargetPath $exePath -WorkingDirectory $InstallDir
Ensure-Shortcut -ShortcutPath (Join-Path $startMenuDir "Uninstall.lnk") -TargetPath "powershell.exe" -WorkingDirectory $InstallDir

$desktopLink = Join-Path ([Environment]::GetFolderPath("Desktop")) "VMWV.Modern.lnk"
Ensure-Shortcut -ShortcutPath $desktopLink -TargetPath $exePath -WorkingDirectory $InstallDir

New-Item -Path $uninstallKey -Force | Out-Null
Set-ItemProperty -Path $uninstallKey -Name "DisplayName" -Value $appName
Set-ItemProperty -Path $uninstallKey -Name "DisplayVersion" -Value "1.0.0"
Set-ItemProperty -Path $uninstallKey -Name "Publisher" -Value "VMWV.Modern"
Set-ItemProperty -Path $uninstallKey -Name "InstallLocation" -Value $InstallDir
Set-ItemProperty -Path $uninstallKey -Name "DisplayIcon" -Value "$exePath,0"
Set-ItemProperty -Path $uninstallKey -Name "UninstallString" -Value "powershell.exe -ExecutionPolicy Bypass -File `"$InstallDir\uninstall.ps1`""

if ($LaunchAfterInstall) {
    Write-Log "Launching application"
    Start-Process -FilePath $exePath -WorkingDirectory $InstallDir
} else {
    Write-Log "Install complete. Launch manually from: $exePath"
}

Write-Log "Install complete"

} catch {
    $errMsg = $_.Exception.Message
    $errLine = $_.InvocationInfo.ScriptLineNumber
    Write-Host "[ERROR] $errMsg (line $errLine)" -ForegroundColor Red
    Add-Content -Path $logFile -Value "[ERROR] $errMsg (line $errLine)"
    Add-Content -Path $logFile -Value $_.ScriptStackTrace
    Add-Content -Path $logFile -Value "--- INSTALL FAILED ---"
    exit 1
}
