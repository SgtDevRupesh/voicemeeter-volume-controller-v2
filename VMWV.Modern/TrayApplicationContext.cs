using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using VMWV.Modern.Services;

namespace VMWV.Modern;

internal sealed class TrayApplicationContext : ApplicationContext
{
    private readonly NotifyIcon _notifyIcon;

    private readonly ToolStripMenuItem _restartOnDeviceChangeMenuItem;
    private readonly ToolStripMenuItem _restartOnAnyDeviceChangeMenuItem;
    private readonly ToolStripMenuItem _restartOnResumeMenuItem;
    private readonly ToolStripMenuItem _restartOnAppLaunchMenuItem;

    private readonly ToolStripMenuItem _startWithWindowsMenuItem;
    private readonly ToolStripMenuItem _limitDbGainMenuItem;
    private readonly ToolStripMenuItem _linearScaleMenuItem;
    private readonly ToolStripMenuItem _syncMuteMenuItem;
    private readonly ToolStripMenuItem _rememberVolumeMenuItem;
    private readonly ToolStripMenuItem _preventVolumeSpikesMenuItem;
    private readonly ToolStripMenuItem _crackleFixMenuItem;

    private readonly List<ToolStripMenuItem> _stripBindingItems;
    private readonly List<ToolStripMenuItem> _busBindingItems;

    private readonly PowerShellDeviceCounter _audioDeviceCounter;
    private readonly PowerShellDeviceCounter _anyDeviceCounter;
    private readonly WindowsDeviceChangeWatcher _deviceChangeWatcher;
    private readonly ResumeWatcher _resumeWatcher;
    private readonly SettingsStore _settingsStore;
    private readonly VoicemeeterController _voicemeeterController;
    private readonly StartupTaskService _startupTaskService;
    private readonly SystemAudioService _systemAudioService;
    private readonly System.Threading.Timer _bindingLabelRefreshTimer;

    private readonly CancellationTokenSource _cancellationTokenSource;
    private readonly SemaphoreSlim _deviceChangeGate;
    private readonly object _logLock;
    private readonly string _logFilePath;
    private readonly string _settingsFilePath;

    private AppSettings _settings;

    private int? _lastAudioDeviceCount;
    private int? _lastAnyDeviceCount;
    private int? _lastKnownVolume;
    private DateTimeOffset _lastVolumeTime;
    private bool _bindingLabelsPrimed;
    private bool _disposed;

    public TrayApplicationContext()
    {
        var appDataDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VMWV.Modern"
        );
        _settingsFilePath = Path.Combine(appDataDirectory, "settings.json");
        _logFilePath = Path.Combine(
            appDataDirectory,
            "logs",
            "vmwv-modern.log"
        );

        RotateLogIfExists(_logFilePath);

        _settingsStore = new SettingsStore(_settingsFilePath);
        _cancellationTokenSource = new CancellationTokenSource();
        _deviceChangeGate = new SemaphoreSlim(1, 1);
        _logLock = new object();
        _settings = _settingsStore.Load();
        _lastVolumeTime = DateTimeOffset.Now;

        _stripBindingItems = new List<ToolStripMenuItem>();
        _busBindingItems = new List<ToolStripMenuItem>();

        _startupTaskService = new StartupTaskService(Log);
        _audioDeviceCounter = new PowerShellDeviceCounter(
            getDeviceCountCommand: "(Get-PnpDevice -Class AudioEndpoint -PresentOnly -ErrorAction SilentlyContinue | Measure-Object).Count"
        );
        _anyDeviceCounter = new PowerShellDeviceCounter(
            getDeviceCountCommand: BuildAnyDeviceCountCommand(_settings.DeviceBlacklist)
        );
        _systemAudioService = new SystemAudioService(TimeSpan.FromMilliseconds(Math.Max(50, _settings.PollingRateMs)));
        _systemAudioService.AudioChanged += OnSystemAudioChanged;
        _bindingLabelRefreshTimer = new System.Threading.Timer(
            _ => RefreshBindingLabelsBackground(),
            null,
            TimeSpan.FromSeconds(2),
            TimeSpan.FromSeconds(5)
        );

        _voicemeeterController = new VoicemeeterController(Log);

        _deviceChangeWatcher = new WindowsDeviceChangeWatcher();
        _deviceChangeWatcher.DeviceChanged += OnDeviceChanged;
        _resumeWatcher = new ResumeWatcher();
        _resumeWatcher.Resumed += OnSystemResumed;

        _restartOnDeviceChangeMenuItem = CreateToggleMenuItem(
            "Audio Devices Change",
            _settings.RestartAudioEngineOnDeviceChange,
            (value) =>
            {
                _settings.RestartAudioEngineOnDeviceChange = value;
                SaveSettings();
            }
        );
        _restartOnAnyDeviceChangeMenuItem = CreateToggleMenuItem(
            "Any Device Change",
            _settings.RestartAudioEngineOnAnyDeviceChange,
            (value) =>
            {
                _settings.RestartAudioEngineOnAnyDeviceChange = value;
                SaveSettings();
            }
        );
        _restartOnResumeMenuItem = CreateToggleMenuItem(
            "Resume From Standby",
            _settings.RestartAudioEngineOnResume,
            (value) =>
            {
                _settings.RestartAudioEngineOnResume = value;
                SaveSettings();
            }
        );
        _restartOnAppLaunchMenuItem = CreateToggleMenuItem(
            "App Launch",
            _settings.RestartAudioEngineOnAppLaunch,
            (value) =>
            {
                _settings.RestartAudioEngineOnAppLaunch = value;
                SaveSettings();
            }
        );

        _startWithWindowsMenuItem = CreateToggleMenuItem(
            "Automatically Start With Windows",
            _settings.StartWithWindows,
            (value) =>
            {
                _settings.StartWithWindows = value;
                _startupTaskService.SetEnabled(value);
                SaveSettings();
            }
        );
        _limitDbGainMenuItem = CreateToggleMenuItem(
            "Limit Max Gain To 0dB",
            _settings.LimitDbGainTo0,
            (value) =>
            {
                _settings.LimitDbGainTo0 = value;
                SaveSettings();
            }
        );
        _linearScaleMenuItem = CreateToggleMenuItem(
            "Use Linear Volume Scaling",
            _settings.LinearVolumeScale,
            (value) =>
            {
                _settings.LinearVolumeScale = value;
                SaveSettings();
            }
        );
        _syncMuteMenuItem = CreateToggleMenuItem(
            "Sync Mute",
            _settings.SyncMute,
            (value) =>
            {
                _settings.SyncMute = value;
                SaveSettings();
            }
        );
        _rememberVolumeMenuItem = CreateToggleMenuItem(
            "Restore Volume At Launch",
            _settings.RememberVolume,
            (value) =>
            {
                _settings.RememberVolume = value;
                if (value)
                {
                    var currentVolume = _systemAudioService.GetVolumePercent();
                    if (currentVolume.HasValue)
                    {
                        _settings.InitialVolume = currentVolume.Value;
                    }
                }
                SaveSettings();
            }
        );
        _preventVolumeSpikesMenuItem = CreateToggleMenuItem(
            "Prevent 100% Volume Spikes",
            _settings.ApplyVolumeFix,
            (value) =>
            {
                _settings.ApplyVolumeFix = value;
                SaveSettings();
            }
        );
        _crackleFixMenuItem = CreateToggleMenuItem(
            "Apply Crackle Fix (USB Interfaces)",
            _settings.ApplyCrackleFix,
            (value) =>
            {
                _settings.ApplyCrackleFix = value;
                if (value)
                {
                    ProcessTweaksService.ApplyAudiodgTweak(_settings.AudiodgPriority, _settings.AudiodgAffinity, Log);
                }
                else
                {
                    ProcessTweaksService.ResetAudiodgTweak(Log);
                }
                SaveSettings();
            }
        );

        var bindingsMenu = new ToolStripMenuItem("Bind Windows Volume To...");
        bindingsMenu.DropDownItems.Add(new ToolStripMenuItem("INPUTS") { Enabled = false });
        for (var i = 0; i <= 7; i++)
        {
            var index = i;
            var item = CreateToggleMenuItem($"Input Strip {index}", _settings.BoundStrips.Contains(index), (value) =>
            {
                if (value)
                {
                    _settings.BoundStrips.Add(index);
                }
                else
                {
                    _settings.BoundStrips.Remove(index);
                }
                SaveSettings();
            });
            _stripBindingItems.Add(item);
            bindingsMenu.DropDownItems.Add(item);
        }

        bindingsMenu.DropDownItems.Add(new ToolStripSeparator());
        bindingsMenu.DropDownItems.Add(new ToolStripMenuItem("OUTPUTS") { Enabled = false });
        for (var i = 0; i <= 7; i++)
        {
            var index = i;
            var item = CreateToggleMenuItem($"Output Bus {index}", _settings.BoundBuses.Contains(index), (value) =>
            {
                if (value)
                {
                    _settings.BoundBuses.Add(index);
                }
                else
                {
                    _settings.BoundBuses.Remove(index);
                }
                SaveSettings();
            });
            _busBindingItems.Add(item);
            bindingsMenu.DropDownItems.Add(item);
        }

        bindingsMenu.DropDownOpening += (_, _) => PrimeBindingLabelsForFirstOpen();

        var restartsMenu = new ToolStripMenuItem("Restart Audio Engine On...");
        restartsMenu.DropDownItems.Add(_restartOnDeviceChangeMenuItem);
        restartsMenu.DropDownItems.Add(_restartOnAnyDeviceChangeMenuItem);
        restartsMenu.DropDownItems.Add(_restartOnResumeMenuItem);
        restartsMenu.DropDownItems.Add(_restartOnAppLaunchMenuItem);

        var settingsMenu = new ToolStripMenuItem("Settings");
        settingsMenu.DropDownItems.Add(new ToolStripMenuItem("SETTINGS") { Enabled = false });
        settingsMenu.DropDownItems.Add(_startWithWindowsMenuItem);
        settingsMenu.DropDownItems.Add(_limitDbGainMenuItem);
        settingsMenu.DropDownItems.Add(_linearScaleMenuItem);
        settingsMenu.DropDownItems.Add(_syncMuteMenuItem);
        settingsMenu.DropDownItems.Add(new ToolStripSeparator());
        settingsMenu.DropDownItems.Add(new ToolStripMenuItem("PATCHES AND WORKAROUNDS") { Enabled = false });
        settingsMenu.DropDownItems.Add(_rememberVolumeMenuItem);
        settingsMenu.DropDownItems.Add(_preventVolumeSpikesMenuItem);
        settingsMenu.DropDownItems.Add(_crackleFixMenuItem);

        var voicemeeterMenu = new ToolStripMenuItem("VOICEMEETER");
        voicemeeterMenu.DropDownItems.Add(new ToolStripMenuItem("Show Voicemeeter", null, (_, _) => ShowVoicemeeter()));
        voicemeeterMenu.DropDownItems.Add(new ToolStripMenuItem("Restart Voicemeeter", null, (_, _) => RestartVoicemeeter()));
        voicemeeterMenu.DropDownItems.Add(new ToolStripMenuItem("Restart Audio Engine", null, (_, _) => RestartAudioEngine("User Input")));

        var supportMenu = new ToolStripMenuItem("SUPPORT");
        supportMenu.DropDownItems.Add(new ToolStripMenuItem("Open Application Folder", null, (_, _) => OpenApplicationFolder()));
        supportMenu.DropDownItems.Add(new ToolStripMenuItem("Open Logs Folder", null, (_, _) => OpenLogsFolder()));

        var menu = new ContextMenuStrip();
        menu.Items.Add(new ToolStripMenuItem("VMWV.Modern") { Enabled = false });
        menu.Items.Add(bindingsMenu);
        menu.Items.Add(restartsMenu);
        menu.Items.Add(settingsMenu);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(voicemeeterMenu);
        menu.Items.Add(supportMenu);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem("Exit", null, (_, _) => ExitThread()));

        _notifyIcon = new NotifyIcon
        {
            Text = "Voicemeeter Windows Volume Modern",
            Icon = (Icon)SystemIcons.Application.Clone(),
            Visible = true,
            ContextMenuStrip = menu,
        };
        _notifyIcon.DoubleClick += (_, _) => RestartAudioEngine("User Input");

        Log("VMWV.Modern started.");

        if (_settings.StartWithWindows)
        {
            var startupEnabled = _startupTaskService.SetEnabled(true);
            if (!startupEnabled)
            {
                _settings.StartWithWindows = false;
                _startWithWindowsMenuItem.Checked = false;
                SaveSettings();
            }
        }
        if (_settings.ApplyCrackleFix)
        {
            ProcessTweaksService.ApplyAudiodgTweak(_settings.AudiodgPriority, _settings.AudiodgAffinity, Log);
        }

        _systemAudioService.Start();
        SetInitialVolumeIfEnabled();

        _ = RefreshDeviceCountsAsync(triggerRestartIfChanged: false, reason: "Startup");

        if (_settings.RestartAudioEngineOnAppLaunch)
        {
            RestartAudioEngine("App Launch");
        }
    }

    protected override void ExitThreadCore()
    {
        Dispose();
        base.ExitThreadCore();
    }

    protected override void Dispose(bool disposing)
    {
        if (_disposed)
        {
            return;
        }

        if (disposing)
        {
            _deviceChangeWatcher.DeviceChanged -= OnDeviceChanged;
            _deviceChangeWatcher.Dispose();
            _resumeWatcher.Resumed -= OnSystemResumed;
            _resumeWatcher.Dispose();

            _systemAudioService.AudioChanged -= OnSystemAudioChanged;
            _systemAudioService.Dispose();
            _bindingLabelRefreshTimer.Dispose();

            _cancellationTokenSource.Cancel();
            _cancellationTokenSource.Dispose();
            _deviceChangeGate.Dispose();

            _voicemeeterController.Dispose();

            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
        }

        _disposed = true;
        base.Dispose(disposing);
    }

    private async void OnDeviceChanged(object? sender, EventArgs e)
    {
        Log("Device change event received from Windows.");
        await RefreshDeviceCountsAsync(triggerRestartIfChanged: true, reason: "Device Change");
    }

    private async void OnSystemResumed(object? sender, EventArgs e)
    {
        Log("System resume detected.");
        if (_settings.RestartAudioEngineOnResume)
        {
            await Task.Delay(1000);
            RestartAudioEngine("Resume From Standby");

            if (_settings.ApplyCrackleFix)
            {
                await Task.Delay(3000);
                ProcessTweaksService.ApplyAudiodgTweak(_settings.AudiodgPriority, _settings.AudiodgAffinity, Log);
            }
        }
    }

    private async Task RefreshDeviceCountsAsync(bool triggerRestartIfChanged, string reason)
    {
        if (_disposed)
        {
            return;
        }

        var lockAcquired = false;
        try
        {
            await _deviceChangeGate.WaitAsync(_cancellationTokenSource.Token);
            lockAcquired = true;

            var audioCount = await _audioDeviceCounter.TryGetCurrentCountAsync(_cancellationTokenSource.Token);
            var anyCount = await _anyDeviceCounter.TryGetCurrentCountAsync(_cancellationTokenSource.Token);

            if (!audioCount.HasValue && !anyCount.HasValue)
            {
                return;
            }

            if (audioCount.HasValue && !_lastAudioDeviceCount.HasValue)
            {
                _lastAudioDeviceCount = audioCount.Value;
                Log($"Initial audio device count detected: {_lastAudioDeviceCount.Value}.");
            }
            if (anyCount.HasValue && !_lastAnyDeviceCount.HasValue)
            {
                _lastAnyDeviceCount = anyCount.Value;
                Log($"Initial any-device count detected: {_lastAnyDeviceCount.Value}.");
            }

            var restartByAudioDevice = false;
            if (audioCount.HasValue && _lastAudioDeviceCount.HasValue && _lastAudioDeviceCount.Value != audioCount.Value)
            {
                Log($"Audio device count changed ({_lastAudioDeviceCount.Value} => {audioCount.Value}).");
                restartByAudioDevice = _settings.RestartAudioEngineOnDeviceChange && audioCount.Value > 0;
                _lastAudioDeviceCount = audioCount.Value;
            }

            var restartByAnyDevice = false;
            if (anyCount.HasValue && _lastAnyDeviceCount.HasValue && _lastAnyDeviceCount.Value != anyCount.Value)
            {
                Log($"Any-device count changed ({_lastAnyDeviceCount.Value} => {anyCount.Value}).");
                restartByAnyDevice = _settings.RestartAudioEngineOnAnyDeviceChange && anyCount.Value > 0;
                _lastAnyDeviceCount = anyCount.Value;
            }

            if (!triggerRestartIfChanged)
            {
                return;
            }

            if (restartByAudioDevice || restartByAnyDevice)
            {
                RestartAudioEngine(reason);
            }
        }
        catch (OperationCanceledException)
        {
            // ignore during shutdown
        }
        finally
        {
            if (lockAcquired)
            {
                _deviceChangeGate.Release();
            }
        }
    }

    private void OnSystemAudioChanged(object? sender, SystemAudioChangedEventArgs e)
    {
        Log($"Audio changed: volume={e.NewVolume}% mute={e.NewMute}");

        if (_settings.ApplyVolumeFix && e.NewVolume == 100)
        {
            var elapsed = e.When - _lastVolumeTime;
            if (elapsed.TotalMilliseconds >= 1000 && _lastKnownVolume.HasValue && _settings.InitialVolume != 100)
            {
                _systemAudioService.SetVolumePercent(_lastKnownVolume.Value);
                Log($"Driver anomaly detected: volume jumped to 100 from {_lastKnownVolume.Value}. Reverting.");
                return;
            }
        }

        _lastKnownVolume = e.NewVolume;
        _lastVolumeTime = e.When;

        if (_settings.RememberVolume)
        {
            _settings.InitialVolume = e.NewVolume;
            SaveSettings();
        }

        PropagateVolumeToVoicemeeter(e.NewVolume);

        if (_settings.SyncMute && e.OldMute.HasValue && e.OldMute.Value != e.NewMute)
        {
            PropagateMuteToVoicemeeter(e.NewMute);
        }
    }

    private void RestartAudioEngine(string reason)
    {
        var result = _voicemeeterController.TryRestartAudioEngine(reason);
        if (result)
        {
            Log($"Audio engine restart command accepted. Reason: {reason}");
            return;
        }

        Log($"Audio engine restart failed. Check log for Voicemeeter API details. Reason: {reason}");
    }

    private void PropagateVolumeToVoicemeeter(int windowsVolume)
    {
        var gainMax = _settings.LimitDbGainTo0 ? 0 : _settings.GainMax;
        var gain = _settings.LinearVolumeScale
            ? ConvertVolumeToGainLinear(windowsVolume, _settings.GainMin, gainMax)
            : ConvertVolumeToGainLogarithmic(windowsVolume, _settings.GainMin, gainMax);

        var boundStrips = _settings.BoundStrips;
        var boundBuses = _settings.BoundBuses;

        Log($"Propagating gain={gain:0.0}dB (vol={windowsVolume}%) to strips=[{string.Join(",", boundStrips)}] buses=[{string.Join(",", boundBuses)}]");

        foreach (var strip in boundStrips)
        {
            var ok = _voicemeeterController.TrySetGain("Strip", strip, gain);
            Log($"Set Strip[{strip}].Gain={gain:0.0} -> {(ok ? "OK" : "FAILED")}");
        }
        foreach (var bus in boundBuses)
        {
            var ok = _voicemeeterController.TrySetGain("Bus", bus, gain);
            Log($"Set Bus[{bus}].Gain={gain:0.0} -> {(ok ? "OK" : "FAILED")}");
        }
    }

    private void PropagateMuteToVoicemeeter(bool mute)
    {
        foreach (var strip in _settings.BoundStrips)
        {
            _voicemeeterController.TrySetMute("Strip", strip, mute);
        }
        foreach (var bus in _settings.BoundBuses)
        {
            _voicemeeterController.TrySetMute("Bus", bus, mute);
        }
    }

    private void SetInitialVolumeIfEnabled()
    {
        if (!_settings.RememberVolume || !_settings.InitialVolume.HasValue)
        {
            return;
        }

        _systemAudioService.SetVolumePercent(_settings.InitialVolume.Value);
        Log($"Set initial volume to {_settings.InitialVolume.Value}%.");
    }

    private ToolStripMenuItem CreateToggleMenuItem(string title, bool initialValue, Action<bool> onChanged)
    {
        var item = new ToolStripMenuItem(title)
        {
            CheckOnClick = true,
            Checked = initialValue,
        };

        item.CheckedChanged += (_, _) => onChanged(item.Checked);
        return item;
    }

    private void RefreshBindingLabelsBackground()
    {
        if (_disposed)
        {
            return;
        }

        // Build the list of all parameters to read in one Login/Logout session
        var parameters = new List<string>(_stripBindingItems.Count * 2 + _busBindingItems.Count * 2);
        for (var i = 0; i < _stripBindingItems.Count; i++)
        {
            parameters.Add($"Strip[{i}].Label");
            parameters.Add($"Strip[{i}].device.name");
        }
        for (var i = 0; i < _busBindingItems.Count; i++)
        {
            parameters.Add($"Bus[{i}].Label");
            parameters.Add($"Bus[{i}].device.name");
        }

        var values = _voicemeeterController.TryGetParameterStrings(parameters);

        if (_disposed)
        {
            return;
        }

        var stripLabels = new string[_stripBindingItems.Count];
        for (var i = 0; i < _stripBindingItems.Count; i++)
        {
            values.TryGetValue($"Strip[{i}].Label", out var label);
            values.TryGetValue($"Strip[{i}].device.name", out var deviceName);
            var baseLabel = string.IsNullOrWhiteSpace(label) ? $"Input Strip {i}" : label.Trim();
            stripLabels[i] = string.IsNullOrWhiteSpace(deviceName) ? baseLabel : $"{baseLabel} : <{deviceName.Trim()}>";
        }

        var busLabels = new string[_busBindingItems.Count];
        for (var i = 0; i < _busBindingItems.Count; i++)
        {
            values.TryGetValue($"Bus[{i}].Label", out var label);
            values.TryGetValue($"Bus[{i}].device.name", out var deviceName);
            var baseLabel = string.IsNullOrWhiteSpace(label) ? $"Output Bus {i}" : label.Trim();
            busLabels[i] = string.IsNullOrWhiteSpace(deviceName) ? baseLabel : $"{baseLabel} : <{deviceName.Trim()}>";
        }

        // Marshal text updates back to the UI thread
        try
        {
            _notifyIcon.ContextMenuStrip?.BeginInvoke(() =>
            {
                if (_disposed)
                {
                    return;
                }

                for (var i = 0; i < _stripBindingItems.Count; i++)
                {
                    _stripBindingItems[i].Text = stripLabels[i];
                }
                for (var i = 0; i < _busBindingItems.Count; i++)
                {
                    _busBindingItems[i].Text = busLabels[i];
                }
            });
        }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { }
    }

    private void SaveSettings()
    {
        _settingsStore.Save(_settings);
    }

    private static double ConvertVolumeToGainLinear(int windowsVolume, int gainMin, int gainMax)
    {
        var gain = (windowsVolume * (gainMax - gainMin)) / 100.0 + gainMin;
        return Math.Round(gain, 1);
    }

    private static double ConvertVolumeToGainLogarithmic(int windowsVolume, int gainMin, int gainMax)
    {
        var amp = windowsVolume > 0 ? Math.Log10(windowsVolume / 100.0) : -1000;
        var gain = Math.Max(20 * amp + gainMax, gainMin);
        return Math.Round(gain, 1);
    }

    private string BuildAnyDeviceCountCommand(IEnumerable<string> blacklist)
    {
        var names = blacklist.Where(x => !string.IsNullOrWhiteSpace(x)).Select(Regex.Escape).ToArray();
        var blacklistRegex = names.Length == 0 ? "$^" : "^(" + string.Join("|", names) + ")$";

        return "(Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FriendlyName | " +
               "Where-Object { $_ -and $_ -notmatch '" + blacklistRegex + "' } | Measure-Object).Count";
    }

    private void ShowVoicemeeter()
    {
        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "VB", "Voicemeeter", "voicemeeter8.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "VB", "Voicemeeter", "voicemeeterpro.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "VB", "Voicemeeter", "voicemeeter.exe"),
        };

        var path = candidates.FirstOrDefault(File.Exists);
        if (path == null)
        {
            Log("Voicemeeter executable not found.");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true,
        });
    }

    private void RestartVoicemeeter()
    {
        try
        {
            var existing = Process.GetProcesses()
                .Where(p => p.ProcessName.StartsWith("voicemeeter", StringComparison.OrdinalIgnoreCase))
                .ToList();

            string? relaunchPath = null;
            foreach (var process in existing)
            {
                try
                {
                    relaunchPath ??= process.MainModule?.FileName;
                    process.Kill();
                }
                catch
                {
                    // ignore inaccessible process handles
                }
            }

            if (!string.IsNullOrWhiteSpace(relaunchPath) && File.Exists(relaunchPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = relaunchPath,
                    UseShellExecute = true,
                });
            }
            else
            {
                ShowVoicemeeter();
            }
        }
        catch (Exception ex)
        {
            Log("Failed to restart Voicemeeter: " + ex.Message);
        }
    }

    private void PrimeBindingLabelsForFirstOpen()
    {
        if (_disposed || _bindingLabelsPrimed)
        {
            return;
        }

        var parameters = new List<string>(_stripBindingItems.Count * 2 + _busBindingItems.Count * 2);
        for (var i = 0; i < _stripBindingItems.Count; i++)
        {
            parameters.Add($"Strip[{i}].Label");
            parameters.Add($"Strip[{i}].device.name");
        }
        for (var i = 0; i < _busBindingItems.Count; i++)
        {
            parameters.Add($"Bus[{i}].Label");
            parameters.Add($"Bus[{i}].device.name");
        }

        var values = _voicemeeterController.TryGetParameterStrings(parameters);

        for (var i = 0; i < _stripBindingItems.Count; i++)
        {
            values.TryGetValue($"Strip[{i}].Label", out var label);
            values.TryGetValue($"Strip[{i}].device.name", out var deviceName);
            var baseLabel = string.IsNullOrWhiteSpace(label) ? $"Input Strip {i}" : label.Trim();
            _stripBindingItems[i].Text = string.IsNullOrWhiteSpace(deviceName) ? baseLabel : $"{baseLabel} : <{deviceName.Trim()}>";
        }

        for (var i = 0; i < _busBindingItems.Count; i++)
        {
            values.TryGetValue($"Bus[{i}].Label", out var label);
            values.TryGetValue($"Bus[{i}].device.name", out var deviceName);
            var baseLabel = string.IsNullOrWhiteSpace(label) ? $"Output Bus {i}" : label.Trim();
            _busBindingItems[i].Text = string.IsNullOrWhiteSpace(deviceName) ? baseLabel : $"{baseLabel} : <{deviceName.Trim()}>";
        }

        _bindingLabelsPrimed = true;
    }

    private void OpenApplicationFolder()
    {
        var path = AppContext.BaseDirectory;
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = path,
            UseShellExecute = true,
        });
    }

    private void OpenLogsFolder()
    {
        var logDir = Path.GetDirectoryName(_logFilePath);
        if (string.IsNullOrWhiteSpace(logDir))
        {
            return;
        }

        Directory.CreateDirectory(logDir);
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = logDir,
            UseShellExecute = true,
        });
    }

    private static void RotateLogIfExists(string logFilePath)
    {
        if (!File.Exists(logFilePath))
        {
            return;
        }

        try
        {
            var logDir = Path.GetDirectoryName(logFilePath)!;
            var stamp = File.GetLastWriteTime(logFilePath).ToString("yyyy-MM-dd_HH-mm-ss");
            var archivedPath = Path.Combine(logDir, $"vmwv-modern-{stamp}.log");

            // Avoid collision if two instances start within the same second
            if (File.Exists(archivedPath))
            {
                archivedPath = Path.Combine(logDir, $"vmwv-modern-{stamp}-{Guid.NewGuid():N[..6]}.log");
            }

            File.Move(logFilePath, archivedPath);
        }
        catch
        {
            // Non-fatal; proceed with normal logging
        }
    }

    private void Log(string message)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var line = $"[{timestamp}] {message}";
        Debug.WriteLine(line);

        lock (_logLock)
        {
            try
            {
                var logDir = Path.GetDirectoryName(_logFilePath);
                if (!string.IsNullOrWhiteSpace(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                File.AppendAllText(_logFilePath, line + Environment.NewLine);
            }
            catch
            {
                // Logging must never destabilize the tray process.
            }
        }
    }
}
