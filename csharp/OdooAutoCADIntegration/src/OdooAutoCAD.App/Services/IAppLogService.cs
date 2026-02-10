using System.Collections.ObjectModel;

namespace OdooAutoCAD.App.Services;

public enum AppLogLevel { Debug, Info, Warning, Error }

public record AppLogEntry(DateTime Timestamp, string Source, AppLogLevel Level, string Message);

public interface IAppLogService
{
    ObservableCollection<AppLogEntry> LogEntries { get; }
    void Log(string message, string source = "App", AppLogLevel level = AppLogLevel.Info);
    void Clear();
}
