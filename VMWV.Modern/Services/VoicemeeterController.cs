using System.Runtime.InteropServices;
using System.Text;

namespace VMWV.Modern.Services;

internal sealed class VoicemeeterController : IDisposable
{
    private readonly Action<string> _log;
    private readonly object _apiLock = new();

    private nint _libraryHandle;
    private bool _sessionOpen;
    private bool _initialized;
    private bool _disposed;

    private VBVMR_LoginDelegate? _login;
    private VBVMR_LogoutDelegate? _logout;
    private VBVMR_SetParametersDelegate? _setParameters;
    private VBVMR_GetParameterStringDelegate? _getParameterString;

    public VoicemeeterController(Action<string> log)
    {
        _log = log;
    }

    public bool TryRestartAudioEngine(string reason)
    {
        return TryExecuteScript("Command.Restart = 1;", $"restart ({reason})");
    }

    public bool TrySetGain(string targetType, int index, double gain)
    {
        var gainStr = gain.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture);
        var script = $"{targetType}[{index}].Gain={gainStr};";
        return TryExecuteScript(script, $"set {targetType}[{index}] gain");
    }

    public bool TrySetMute(string targetType, int index, bool mute)
    {
        var script = $"{targetType}[{index}].Mute={(mute ? 1 : 0)};";
        return TryExecuteScript(script, $"set {targetType}[{index}] mute");
    }

    public Dictionary<string, string> TryGetParameterStrings(IEnumerable<string> parameters)
    {
        var result = new Dictionary<string, string>();

        if (_disposed || !EnsureSession())
        {
            return result;
        }

        lock (_apiLock)
        {
            foreach (var parameter in parameters)
            {
                var buffer = new StringBuilder(1024);
                var code = _getParameterString!(parameter, buffer);
                if (code >= 0)
                {
                    result[parameter] = buffer.ToString();
                }
            }
        }

        return result;
    }

    public bool TryExecuteScript(string script, string operation)
    {
        if (_disposed || !EnsureSession())
        {
            return false;
        }

        lock (_apiLock)
        {
            var commandResult = _setParameters!(script);
            if (commandResult < 0)
            {
                _log($"Voicemeeter command failed with code {commandResult} during {operation}. Script: {script}");
                return false;
            }

            return true;
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        if (_sessionOpen && _logout != null)
        {
            _logout();
            _sessionOpen = false;
        }

        if (_libraryHandle != nint.Zero)
        {
            NativeLibrary.Free(_libraryHandle);
            _libraryHandle = nint.Zero;
        }
    }

    /// <summary>
    /// Ensures the DLL is loaded and a Voicemeeter session is open.
    /// Logs in once and keeps the session alive for the lifetime of this object.
    /// </summary>
    private bool EnsureSession()
    {
        if (_sessionOpen)
        {
            return true;
        }

        if (!EnsureInitialized())
        {
            return false;
        }

        var loginResult = _login!();
        if (loginResult < 0)
        {
            _log($"Voicemeeter login failed with code {loginResult}.");
            return false;
        }

        _sessionOpen = true;
        _log($"Voicemeeter session opened (login code {loginResult}).");
        return true;
    }

    private bool EnsureInitialized()
    {
        if (_initialized)
        {
            return _libraryHandle != nint.Zero;
        }

        _initialized = true;

        var dllPath = ResolveVoicemeeterRemoteDllPath();
        if (string.IsNullOrWhiteSpace(dllPath) || !File.Exists(dllPath))
        {
            _log("VoicemeeterRemote64.dll not found. Set VOICEMEETER_REMOTE_DLL or install Voicemeeter.");
            return false;
        }

        if (!NativeLibrary.TryLoad(dllPath, out _libraryHandle) || _libraryHandle == nint.Zero)
        {
            _log($"Failed to load Voicemeeter remote DLL from: {dllPath}");
            return false;
        }

        if (!TryLoadExports(_libraryHandle))
        {
            NativeLibrary.Free(_libraryHandle);
            _libraryHandle = nint.Zero;
            return false;
        }

        _log($"Loaded Voicemeeter remote DLL from: {dllPath}");
        return true;
    }

    private bool TryLoadExports(nint handle)
    {
        if (!TryGetDelegate(handle, "VBVMR_Login", out _login))
        {
            _log("Voicemeeter export missing: VBVMR_Login");
            return false;
        }

        if (!TryGetDelegate(handle, "VBVMR_Logout", out _logout))
        {
            _log("Voicemeeter export missing: VBVMR_Logout");
            return false;
        }

        if (!TryGetDelegate(handle, "VBVMR_SetParameters", out _setParameters))
        {
            _log("Voicemeeter export missing: VBVMR_SetParameters");
            return false;
        }

        if (!TryGetDelegate(handle, "VBVMR_GetParameterStringA", out _getParameterString))
        {
            _log("Voicemeeter export missing: VBVMR_GetParameterStringA");
            return false;
        }

        return true;
    }

    private static bool TryGetDelegate<TDelegate>(nint handle, string exportName, out TDelegate? del)
        where TDelegate : Delegate
    {
        del = null;
        if (!NativeLibrary.TryGetExport(handle, exportName, out var proc) || proc == nint.Zero)
        {
            return false;
        }

        del = Marshal.GetDelegateForFunctionPointer<TDelegate>(proc);
        return true;
    }

    private static string? ResolveVoicemeeterRemoteDllPath()
    {
        var envOverride = Environment.GetEnvironmentVariable("VOICEMEETER_REMOTE_DLL");
        if (!string.IsNullOrWhiteSpace(envOverride) && File.Exists(envOverride))
        {
            return envOverride;
        }

        var candidates = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "VB", "Voicemeeter", "VoicemeeterRemote64.dll"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "VB", "Voicemeeter", "VoicemeeterRemote64.dll"),
            Path.Combine(AppContext.BaseDirectory, "VoicemeeterRemote64.dll"),
        };

        return candidates.FirstOrDefault(File.Exists);
    }

    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate int VBVMR_LoginDelegate();

    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate int VBVMR_LogoutDelegate();

    [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private delegate int VBVMR_SetParametersDelegate([MarshalAs(UnmanagedType.LPStr)] string script);

    [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    private delegate int VBVMR_GetParameterStringDelegate([MarshalAs(UnmanagedType.LPStr)] string parameter, StringBuilder value);
}
