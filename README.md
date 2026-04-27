# VVC — Voicemeeter Volume Controller

A lightweight Windows system tray application that bridges Windows system volume to [Voicemeeter](https://vb-audio.com/Voicemeeter/). When you move the Windows volume slider or press your keyboard's mute button, VVC forwards those changes to the Voicemeeter input strips and output buses of your choice.

Built with .NET 8 and WinForms. No runtime installation required — ships as a self-contained single-file executable.

---

## Features

### Volume & Mute Sync
- Polls Windows default audio endpoint and mirrors volume changes to bound Voicemeeter strips/buses
- Supports **linear** or **logarithmic (dB)** volume scaling
- Optional **Sync Mute** — mute/unmute state follows the Windows mute button
- Optional **Limit Max Gain to 0 dB** — prevents clipping from over-amplification
- **Restore Volume At Launch** — saves and restores the last known volume on next startup
- **Prevent 100% Volume Spikes** — guards against sudden full-volume jumps on device changes

### Strip & Bus Binding
- Bind Windows volume to any combination of up to 8 Voicemeeter **Input Strips** and/or 8 **Output Buses**
- Binding labels show the live Voicemeeter channel names (refreshed every 5 seconds)

### Automatic Audio Engine Restart
Trigger a Voicemeeter audio engine restart automatically on:
- **Audio device change** — monitors audio endpoint plug/unplug events via `WM_DEVICECHANGE`
- **Any device change** — watches all PnP device arrivals/removals (configurable blacklist filters noise)
- **Resume from standby** — reacts to `PowerModes.Resume` system events
- **App launch** — restarts once on startup

### Crackle Fix (USB Audio Interfaces)
Adjusts `audiodg.exe` process priority and CPU affinity via PowerShell to reduce audio crackling on USB audio devices. Priority and affinity values are configurable in `settings.json`.

### Startup & System Integration
- **Start With Windows** — registers/removes a `HKCU\...\Run` registry entry so VVC launches at login
- Settings are persisted to `%LOCALAPPDATA%\VVC\settings.json`
- Logs are written to `%LOCALAPPDATA%\VVC\logs\vvc.log` (rotated on each launch)

---

## Requirements

- Windows 10 or later (x64)
- [Voicemeeter](https://vb-audio.com/Voicemeeter/) installed (any edition — Banana, Potato, etc.)
  - `VoicemeeterRemote64.dll` must be present (installed alongside Voicemeeter)

---

## Installation

### Option A — Installer EXE (recommended)

Download `Install-VVC.exe` from the [Releases](../../releases) page and run it.

The installer will:
1. Stop any running VVC instance
2. Copy the application to `%APPDATA%\VVC\`
3. Register an uninstaller entry in Add/Remove Programs
4. Optionally launch VVC immediately

### Option B — ZIP Package

Extract `Install-VVC.zip` and run `install.cmd` (or `install.ps1`) from the extracted folder:

```cmd
install.cmd
```

### Uninstall

Run `uninstall.ps1` from the installation folder, or use **Add/Remove Programs** → **VVC**.

---

## Running from Source

```powershell
dotnet run --project .\VMWV.Modern\VMWV.Modern.csproj
```

---

## Building

**Debug / Release build:**

```powershell
dotnet build .\voicemeeter-windows-volume-modern.sln -c Release
```

**Self-contained single-file publish + installer:**

```powershell
powershell -ExecutionPolicy Bypass -File .\build-tools\build.installer.ps1
```

This script:
1. Publishes the app as a self-contained `win-x64` single-file executable to `artifacts\publish\`
2. Packages the payload and installer scripts into `artifacts\installer\Install-VVC.zip`
3. If `iexpress.exe` is available on the system, also produces `artifacts\installer\Install-VVC.exe`

---

## Configuration

Settings are stored at `%LOCALAPPDATA%\VVC\settings.json` and are fully editable by hand. Key options:

| Key | Default | Description |
|---|---|---|
| `PollingRateMs` | `100` | How often (ms) Windows volume is sampled |
| `GainMin` | `-60` | Minimum dB gain sent to Voicemeeter |
| `GainMax` | `12` | Maximum dB gain sent to Voicemeeter |
| `LinearVolumeScale` | `false` | Use linear instead of dB scaling |
| `LimitDbGainTo0` | `false` | Cap gain at 0 dB |
| `SyncMute` | `true` | Mirror Windows mute to Voicemeeter |
| `RememberVolume` | `false` | Restore volume on app launch |
| `ApplyVolumeFix` | `false` | Block 100% volume spikes |
| `ApplyCrackleFix` | `false` | Tweak `audiodg.exe` priority/affinity |
| `AudiodgPriority` | `128` | Priority class for crackle fix |
| `AudiodgAffinity` | `2` | CPU affinity mask for crackle fix |
| `StartWithWindows` | `true` | Auto-start at Windows login |
| `RestartAudioEngineOnDeviceChange` | `false` | Restart engine on audio device change |
| `RestartAudioEngineOnAnyDeviceChange` | `false` | Restart engine on any PnP device change |
| `RestartAudioEngineOnResume` | `false` | Restart engine after resume from sleep |
| `RestartAudioEngineOnAppLaunch` | `false` | Restart engine once on startup |
| `BoundStrips` | `[]` | Voicemeeter input strip indices to bind |
| `BoundBuses` | `[]` | Voicemeeter output bus indices to bind |
| `DeviceBlacklist` | see below | Device name substrings ignored for "any device change" |

Default `DeviceBlacklist` entries: `"Microsoft Streaming Service Proxy"`, `"Volume"`, `"Xvd"`.

---

