using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Extensions.Logging;

namespace OdooAutoCAD.App.Services;

public class AppLogService : IAppLogService
{
    private const int MaxEntries = 2000;
    private readonly ILogger<AppLogService>? _logger;

    public ObservableCollection<AppLogEntry> LogEntries { get; } = new();

    public AppLogService(ILogger<AppLogService>? logger = null)
    {
        _logger = logger;
        Log("Application log service started");
    }

    public void Log(string message, string source = "App", AppLogLevel level = AppLogLevel.Info)
    {
        var entry = new AppLogEntry(DateTime.Now, source, level, message);

        // Forward to Serilog
        ForwardToSerilog(entry);

        // Marshal to UI thread if needed
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(() => AddEntry(entry));
        }
        else
        {
            AddEntry(entry);
        }
    }

    public void Clear()
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(() => LogEntries.Clear());
        }
        else
        {
            LogEntries.Clear();
        }
    }

    private void AddEntry(AppLogEntry entry)
    {
        if (LogEntries.Count >= MaxEntries)
        {
            LogEntries.RemoveAt(0);
        }
        LogEntries.Add(entry);
    }

    private void ForwardToSerilog(AppLogEntry entry)
    {
        var msg = "[{Source}] {Message}";
        switch (entry.Level)
        {
            case AppLogLevel.Debug:
                _logger?.LogDebug(msg, entry.Source, entry.Message);
                break;
            case AppLogLevel.Info:
                _logger?.LogInformation(msg, entry.Source, entry.Message);
                break;
            case AppLogLevel.Warning:
                _logger?.LogWarning(msg, entry.Source, entry.Message);
                break;
            case AppLogLevel.Error:
                _logger?.LogError(msg, entry.Source, entry.Message);
                break;
        }
    }
}
