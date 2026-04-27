namespace VMWV.Modern.Services;

internal sealed class WindowsDeviceChangeWatcher : NativeWindow, IDisposable
{
    private const int WmDeviceChange = 0x0219;
    private const int DbtDevNodesChanged = 0x0007;
    private const int DbtDeviceArrival = 0x8000;
    private const int DbtDeviceRemoveComplete = 0x8004;

    private bool _disposed;

    public event EventHandler? DeviceChanged;

    public WindowsDeviceChangeWatcher()
    {
        CreateHandle(new CreateParams());
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (Handle != nint.Zero)
        {
            DestroyHandle();
        }

        _disposed = true;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmDeviceChange)
        {
            var eventCode = m.WParam.ToInt32();
            if (eventCode == DbtDevNodesChanged || eventCode == DbtDeviceArrival || eventCode == DbtDeviceRemoveComplete)
            {
                DeviceChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        base.WndProc(ref m);
    }
}
