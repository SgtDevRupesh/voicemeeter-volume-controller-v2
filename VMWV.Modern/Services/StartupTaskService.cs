using System.Diagnostics;
using Microsoft.Win32;

namespace VMWV.Modern.Services;

internal sealed class StartupTaskService
{
    private const string RunValueName = "VMWV.Modern";
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private readonly Action<string> _log;

    public StartupTaskService(Action<string> log)
    {
        _log = log;
    }

    public bool SetEnabled(bool enabled)
    {
        if (enabled)
        {
            return Enable();
        }

        return Disable();
    }

    private bool Enable()
    {
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
        {
            _log("Start with Windows skipped: process path unavailable.");
            return false;
        }

        try
        {
            using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true)
                ?? Registry.CurrentUser.CreateSubKey(RunKeyPath, writable: true);

            if (runKey == null)
            {
                _log("Startup registry key unavailable.");
                return false;
            }

            var command = '"' + exePath + '"';
            runKey.SetValue(RunValueName, command, RegistryValueKind.String);
            _log("Enabled startup entry.");
            return true;
        }
        catch (Exception ex)
        {
            _log("Startup enable failed: " + ex.Message);
            return false;
        }
    }

    private bool Disable()
    {
        try
        {
            using var runKey = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
            if (runKey == null)
            {
                _log("Disabled startup entry (registry key missing).");
                return true;
            }

            if (runKey.GetValue(RunValueName) != null)
            {
                runKey.DeleteValue(RunValueName, throwOnMissingValue: false);
            }

            _log("Disabled startup entry.");
            return true;
        }
        catch (Exception ex)
        {
            _log("Startup disable failed: " + ex.Message);
            return false;
        }
    }
}
