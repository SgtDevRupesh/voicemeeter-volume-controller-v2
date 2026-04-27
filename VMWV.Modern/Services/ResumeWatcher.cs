using Microsoft.Win32;

namespace VMWV.Modern.Services;

internal sealed class ResumeWatcher : IDisposable
{
    private bool _disposed;

    public event EventHandler? Resumed;

    public ResumeWatcher()
    {
        SystemEvents.PowerModeChanged += OnPowerModeChanged;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        SystemEvents.PowerModeChanged -= OnPowerModeChanged;
        _disposed = true;
    }

    private void OnPowerModeChanged(object? sender, PowerModeChangedEventArgs e)
    {
        if (e.Mode == PowerModes.Resume)
        {
            Resumed?.Invoke(this, EventArgs.Empty);
        }
    }
}
