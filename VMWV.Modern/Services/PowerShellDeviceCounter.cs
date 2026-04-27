using System.Diagnostics;

namespace VMWV.Modern.Services;

internal sealed class PowerShellDeviceCounter
{
    private readonly string _getDeviceCountCommand;

    public PowerShellDeviceCounter(string getDeviceCountCommand)
    {
        _getDeviceCountCommand = getDeviceCountCommand;
    }

    public async Task<int?> TryGetCurrentCountAsync(CancellationToken cancellationToken)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{_getDeviceCountCommand}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            },
        };

        if (!process.Start())
        {
            return null;
        }

        var output = await process.StandardOutput.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);

        var firstLine = output
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault();

        if (int.TryParse(firstLine, out var count))
        {
            return count;
        }

        return null;
    }
}

internal sealed class DeviceCountChangedEventArgs : EventArgs
{
    public DeviceCountChangedEventArgs(int oldCount, int newCount)
    {
        OldCount = oldCount;
        NewCount = newCount;
    }

    public int OldCount { get; }

    public int NewCount { get; }
}
