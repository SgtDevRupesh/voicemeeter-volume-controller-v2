param(
    [string]$InstallDir = "$env:APPDATA\VVC"
)

$ErrorActionPreference = "Stop"

$appName = "VVC"
$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$appName"

function Stop-RunningProcesses {
    foreach ($name in @("VVC", "VMWV", "VMWV.Modern")) {
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
foreach ($taskName in @("VMWV.Modern", "VVC")) {
    Remove-TaskIfExists -TaskName $taskName
}

if (Test-Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force
}

$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\VVC"
if (Test-Path $startMenuDir) {
    Remove-Item -Path $startMenuDir -Recurse -Force
}

$desktopLink = Join-Path ([Environment]::GetFolderPath("Desktop")) "VVC.lnk"
if (Test-Path $desktopLink) {
    Remove-Item -Path $desktopLink -Force
}

if (Test-Path $uninstallKey) {
    Remove-Item -Path $uninstallKey -Recurse -Force
}

Write-Host "Uninstall complete"
