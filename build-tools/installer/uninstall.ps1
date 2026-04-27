param(
    [string]$InstallDir = "$env:APPDATA\VMWV.Modern"
)

$ErrorActionPreference = "Stop"

$appName = "Voicemeeter Windows Volume Modern"
$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$appName"

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

Stop-RunningProcesses
Remove-TaskIfExists -TaskName "VMWV.Modern"

if (Test-Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force
}

$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\VMWV.Modern"
if (Test-Path $startMenuDir) {
    Remove-Item -Path $startMenuDir -Recurse -Force
}

$desktopLink = Join-Path ([Environment]::GetFolderPath("Desktop")) "VMWV.Modern.lnk"
if (Test-Path $desktopLink) {
    Remove-Item -Path $desktopLink -Force
}

if (Test-Path $uninstallKey) {
    Remove-Item -Path $uninstallKey -Recurse -Force
}

Write-Host "Uninstall complete"
