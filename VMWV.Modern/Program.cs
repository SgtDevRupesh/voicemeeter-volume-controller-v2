namespace VMWV.Modern;

static class Program
{
    private static Mutex? _singleInstanceMutex;

    [STAThread]
    static void Main()
    {
        _singleInstanceMutex = new Mutex(initiallyOwned: true, name: "Global\\VVC.Tray", createdNew: out var createdNew);
        if (!createdNew)
        {
            _singleInstanceMutex.Dispose();
            _singleInstanceMutex = null;
            return;
        }

        var logDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VVC",
            "logs"
        );

        Directory.CreateDirectory(logDirectory);

        Application.ThreadException += (_, args) =>
            WriteFatalLog(logDirectory, "Unhandled UI thread exception", args.Exception);
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            WriteFatalLog(logDirectory, "Unhandled app-domain exception", args.ExceptionObject as Exception);

        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new TrayApplicationContext());
        }
        catch (Exception ex)
        {
            WriteFatalLog(logDirectory, "Fatal startup exception", ex, showMessageBox: true);
        }
        finally
        {
            _singleInstanceMutex?.Dispose();
            _singleInstanceMutex = null;
        }
    }

    private static void WriteFatalLog(string logDirectory, string message, Exception? exception, bool showMessageBox = false)
    {
        try
        {
            Directory.CreateDirectory(logDirectory);
            var logPath = Path.Combine(logDirectory, "fatal-startup.log");
            var text = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}{exception}{Environment.NewLine}{Environment.NewLine}";
            File.AppendAllText(logPath, text);

            if (showMessageBox)
            {
                MessageBox.Show(
                    $"VVC failed to start.\n\n{exception?.Message}\n\nSee: {logPath}",
                    "VVC Startup Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
        catch
        {
            // Avoid recursive failure during crash reporting.
        }
    }
}