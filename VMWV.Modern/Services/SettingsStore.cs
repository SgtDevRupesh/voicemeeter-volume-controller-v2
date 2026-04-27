using System.Text.Json;

namespace VMWV.Modern.Services;

internal sealed class SettingsStore
{
    private readonly string _settingsFilePath;

    public SettingsStore(string settingsFilePath)
    {
        _settingsFilePath = settingsFilePath;
    }

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(_settingsFilePath))
            {
                return new AppSettings();
            }

            var json = File.ReadAllText(_settingsFilePath);
            var parsed = JsonSerializer.Deserialize<AppSettings>(json);
            return parsed ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        var settingsDirectory = Path.GetDirectoryName(_settingsFilePath);
        if (!string.IsNullOrWhiteSpace(settingsDirectory))
        {
            Directory.CreateDirectory(settingsDirectory);
        }

        var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions
        {
            WriteIndented = true,
        });
        File.WriteAllText(_settingsFilePath, json);
    }
}

internal sealed class AppSettings
{
    public int PollingRateMs { get; set; } = 100;

    public int GainMin { get; set; } = -60;

    public int GainMax { get; set; } = 12;

    public bool StartWithWindows { get; set; } = true;

    public bool LimitDbGainTo0 { get; set; }

    public bool LinearVolumeScale { get; set; }

    public bool SyncMute { get; set; } = true;

    public bool RememberVolume { get; set; }

    public bool ApplyVolumeFix { get; set; }

    public bool ApplyCrackleFix { get; set; }

    public bool RestartAudioEngineOnDeviceChange { get; set; }

    public bool RestartAudioEngineOnAnyDeviceChange { get; set; }

    public bool RestartAudioEngineOnResume { get; set; }

    public bool RestartAudioEngineOnAppLaunch { get; set; }

    public int? InitialVolume { get; set; }

    public int AudiodgPriority { get; set; } = 128;

    public int AudiodgAffinity { get; set; } = 2;

    public List<string> DeviceBlacklist { get; set; } = new()
    {
        "Microsoft Streaming Service Proxy",
        "Volume",
        "Xvd",
    };

    public HashSet<int> BoundStrips { get; set; } = new();

    public HashSet<int> BoundBuses { get; set; } = new();
}
