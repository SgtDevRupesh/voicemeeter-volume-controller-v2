# VMWV.Modern (.NET)

This is a vanilla .NET migration scaffold for Voicemeeter Windows Volume.

## What is implemented

- WinForms tray app shell using `NotifyIcon`
- Context menu with:
  - Restart Audio Engine On Device Change (toggle)
  - Restart Audio Engine Now
  - Open Logs Folder
  - Exit
- Native Windows device-change notifications (`WM_DEVICECHANGE`) with on-demand count checks
- Voicemeeter restart via native `VoicemeeterRemote64.dll` API (`VBVMR_Login`, `VBVMR_SetParameters`, `VBVMR_Logout`)
- Settings persistence to `%LOCALAPPDATA%/VMWV.Modern/settings.json`
- Basic logging to:
  - `%LOCALAPPDATA%/VMWV.Modern/logs/vmwv-modern.log`

## What is not implemented yet

- Windows volume sync and mute sync
- Startup task integration

## Run

From this folder:

```powershell
dotnet run --project .\VMWV.Modern\VMWV.Modern.csproj
```

## Build

```powershell
dotnet build .\voicemeeter-windows-volume-modern.sln -c Release
```

## Build Installer (EXE)

This repository includes a no-extra-tools installer pipeline:

```powershell
powershell -ExecutionPolicy Bypass -File .\build-tools\build.installer.ps1
```

Artifacts are written to:

- `artifacts\installer\Install-VMWV-Modern.exe` (IExpress EXE, when available)
- `artifacts\installer\Install-VMWV-Modern.zip` (always generated fallback package)

## Installer Migration Behavior

The installer script attempts to replace the legacy Node-based app by:

- Detecting installed entries named `Voicemeeter Windows Volume` (non-Modern)
- Running its uninstaller when found
- Removing old scheduled task `voicemeeter-windows-volume`
- Installing Modern build into `C:\Program Files\Voicemeeter Windows Volume Modern`

It does not uninstall system Node.js itself. It only migrates the application install.

If you explicitly want to remove Node.js runtime during install, run installer command with:

```cmd
install.cmd -RemoveNodeJsRuntime
```

This is optional and off by default.

## Next migration steps

1. Add device-level filtering/blacklist parity with legacy behavior.
2. Add master-volume sync and mute sync workflow.
3. Add startup task integration and settings UI parity.
4. Add robust reconnect/retry strategy for Voicemeeter API sessions.
