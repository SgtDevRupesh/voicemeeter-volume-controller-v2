using System.Runtime.InteropServices;

namespace VMWV.Modern.Services;

internal sealed class SystemAudioService : IDisposable
{
    private readonly TimeSpan _pollingInterval;
    private readonly System.Windows.Forms.Timer _timer;

    private int? _lastVolume;
    private bool? _lastMute;
    private bool _disposed;

    public event EventHandler<SystemAudioChangedEventArgs>? AudioChanged;

    public SystemAudioService(TimeSpan pollingInterval)
    {
        _pollingInterval = pollingInterval;
        _timer = new System.Windows.Forms.Timer
        {
            Interval = Math.Max(50, (int)_pollingInterval.TotalMilliseconds),
        };
        _timer.Tick += OnTick;
    }

    public void Start()
    {
        _timer.Start();
    }

    public void Stop()
    {
        _timer.Stop();
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _timer.Stop();
        _timer.Tick -= OnTick;
        _timer.Dispose();

        _disposed = true;
    }

    public int? GetVolumePercent()
    {
        using var endpointScope = GetDefaultEndpointVolume();
        var endpoint = endpointScope?.Volume;
        if (endpoint == null)
        {
            return null;
        }

        endpoint.GetMasterVolumeLevelScalar(out var scalar);
        return (int)Math.Round(Math.Clamp(scalar, 0f, 1f) * 100f);
    }

    public bool? GetMute()
    {
        using var endpointScope = GetDefaultEndpointVolume();
        var endpoint = endpointScope?.Volume;
        if (endpoint == null)
        {
            return null;
        }

        endpoint.GetMute(out var mute);
        return mute;
    }

    public bool SetVolumePercent(int volumePercent)
    {
        using var endpointScope = GetDefaultEndpointVolume();
        var endpoint = endpointScope?.Volume;
        if (endpoint == null)
        {
            return false;
        }

        var scalar = Math.Clamp(volumePercent / 100f, 0f, 1f);
        endpoint.SetMasterVolumeLevelScalar(scalar, Guid.Empty);
        return true;
    }

    public bool SetMute(bool mute)
    {
        using var endpointScope = GetDefaultEndpointVolume();
        var endpoint = endpointScope?.Volume;
        if (endpoint == null)
        {
            return false;
        }

        endpoint.SetMute(mute, Guid.Empty);
        return true;
    }

    private void OnTick(object? sender, EventArgs e)
    {
        using var endpointScope = GetDefaultEndpointVolume();
        var endpoint = endpointScope?.Volume;
        if (endpoint == null)
        {
            return;
        }

        endpoint.GetMasterVolumeLevelScalar(out var scalar);
        endpoint.GetMute(out var mute);

        var volume = (int)Math.Round(Math.Clamp(scalar, 0f, 1f) * 100f);

        if (_lastVolume.HasValue && _lastMute.HasValue && _lastVolume.Value == volume && _lastMute.Value == mute)
        {
            return;
        }

        var previousVolume = _lastVolume;
        var previousMute = _lastMute;

        _lastVolume = volume;
        _lastMute = mute;

        AudioChanged?.Invoke(this, new SystemAudioChangedEventArgs(previousVolume, volume, previousMute, mute, DateTimeOffset.Now));
    }

    private static EndpointVolumeScope? GetDefaultEndpointVolume()
    {
        IMMDeviceEnumerator? enumerator = null;
        IMMDevice? device = null;
        object? endpointVolumeObject = null;

        try
        {
            enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
            enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out var endpointDevice);
            device = endpointDevice;

            var iid = typeof(IAudioEndpointVolume).GUID;
            endpointDevice.Activate(ref iid, CLSCTX.ALL, IntPtr.Zero, out endpointVolumeObject);

            return new EndpointVolumeScope((IAudioEndpointVolume)endpointVolumeObject, endpointDevice, enumerator, endpointVolumeObject);
        }
        catch
        {
            ReleaseComObject(endpointVolumeObject);
            ReleaseComObject(device);
            ReleaseComObject(enumerator);
            return null;
        }
    }

    private static void ReleaseComObject(object? comObject)
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }

    private sealed class EndpointVolumeScope : IDisposable
    {
        private readonly object _endpointVolumeObject;
        private readonly IMMDevice _device;
        private readonly IMMDeviceEnumerator _enumerator;
        private bool _disposed;

        public EndpointVolumeScope(
            IAudioEndpointVolume volume,
            IMMDevice device,
            IMMDeviceEnumerator enumerator,
            object endpointVolumeObject)
        {
            Volume = volume;
            _device = device;
            _enumerator = enumerator;
            _endpointVolumeObject = endpointVolumeObject;
        }

        public IAudioEndpointVolume Volume { get; }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            ReleaseComObject(_endpointVolumeObject);
            ReleaseComObject(_device);
            ReleaseComObject(_enumerator);
        }
    }

    private enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
    }

    private enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
    }

    [Flags]
    private enum CLSCTX : uint
    {
        INPROC_SERVER = 0x1,
        INPROC_HANDLER = 0x2,
        LOCAL_SERVER = 0x4,
        INPROC_SERVER16 = 0x8,
        REMOTE_SERVER = 0x10,
        INPROC_HANDLER16 = 0x20,
        RESERVED1 = 0x40,
        RESERVED2 = 0x80,
        RESERVED3 = 0x100,
        RESERVED4 = 0x200,
        NO_CODE_DOWNLOAD = 0x400,
        RESERVED5 = 0x800,
        NO_CUSTOM_MARSHAL = 0x1000,
        ENABLE_CODE_DOWNLOAD = 0x2000,
        NO_FAILURE_LOG = 0x4000,
        DISABLE_AAA = 0x8000,
        ENABLE_AAA = 0x10000,
        FROM_DEFAULT_CONTEXT = 0x20000,
        ACTIVATE_X86_SERVER = 0x40000,
        ACTIVATE_32_BIT_SERVER = ACTIVATE_X86_SERVER,
        ACTIVATE_64_BIT_SERVER = 0x80000,
        ENABLE_CLOAKING = 0x100000,
        APPCONTAINER = 0x400000,
        ACTIVATE_AAA_AS_IU = 0x800000,
        PS_DLL = 0x80000000,
        ALL = INPROC_SERVER | INPROC_HANDLER | LOCAL_SERVER | REMOTE_SERVER,
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    private class MMDeviceEnumeratorComObject
    {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceEnumerator
    {
        int NotImpl1();

        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDevice
    {
        [PreserveSig]
        int Activate(ref Guid iid, CLSCTX dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
    }

    [ComImport]
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IntPtr pNotify);

        int UnregisterControlChangeNotify(IntPtr pNotify);

        int GetChannelCount(out uint pnChannelCount);

        int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);

        int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);

        int GetMasterVolumeLevel(out float pfLevelDB);

        int GetMasterVolumeLevelScalar(out float pfLevel);

        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);

        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);

        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);

        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);

        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);

        int GetMute(out bool pbMute);

        int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);

        int VolumeStepUp(Guid pguidEventContext);

        int VolumeStepDown(Guid pguidEventContext);

        int QueryHardwareSupport(out uint pdwHardwareSupportMask);

        int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }
}

internal sealed class SystemAudioChangedEventArgs : EventArgs
{
    public SystemAudioChangedEventArgs(int? oldVolume, int newVolume, bool? oldMute, bool newMute, DateTimeOffset when)
    {
        OldVolume = oldVolume;
        NewVolume = newVolume;
        OldMute = oldMute;
        NewMute = newMute;
        When = when;
    }

    public int? OldVolume { get; }

    public int NewVolume { get; }

    public bool? OldMute { get; }

    public bool NewMute { get; }

    public DateTimeOffset When { get; }
}
