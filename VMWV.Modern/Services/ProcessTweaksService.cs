using System.Diagnostics;

namespace VMWV.Modern.Services;

internal static class ProcessTweaksService
{
    public const int PriorityNormal = 32;

    public static void ApplyAudiodgTweak(int priority, int affinity, Action<string> log)
    {
        var command = "Get-WmiObject Win32_process -filter 'name = \"audiodg.exe\"' | foreach-object { $_.SetPriority(" + priority + ") }; " +
                      "$process=Get-Process audiodg -ErrorAction SilentlyContinue; if ($process) { $process.ProcessorAffinity=" + affinity + " }";
        RunPowerShell(command, log, "Applied audiodg tweak");
    }

    public static void ResetAudiodgTweak(Action<string> log)
    {
        var command = "Get-WmiObject Win32_process -filter 'name = \"audiodg.exe\"' | foreach-object { $_.SetPriority(" + PriorityNormal + ") }; " +
                      "$process=Get-Process audiodg -ErrorAction SilentlyContinue; if ($process) { $process.ProcessorAffinity=255 }";
        RunPowerShell(command, log, "Reset audiodg tweak");
    }

    private static void RunPowerShell(string script, Action<string> log, string successMessage)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"" + script + "\"",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                },
            };

            process.Start();
            var error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (process.ExitCode == 0)
            {
                log(successMessage + ".");
                return;
            }

            log("audiodg command failed: " + error.Trim());
        }
        catch (Exception ex)
        {
            log("audiodg command threw: " + ex.Message);
        }
    }
}
