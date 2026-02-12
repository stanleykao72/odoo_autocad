using System.Collections.Specialized;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using OdooAutoCAD.App.Services;

namespace OdooAutoCAD.App.Views.Controls;

public partial class LogPanelControl : UserControl
{
    private string _sourceFilter = "All";
    private string _levelFilter = "All";
    private readonly HashSet<string> _knownSources = new() { "All" };

    public LogPanelControl()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is IAppLogService oldService)
        {
            oldService.LogEntries.CollectionChanged -= OnLogEntriesChanged;
        }

        if (e.NewValue is IAppLogService newService)
        {
            newService.LogEntries.CollectionChanged += OnLogEntriesChanged;

            // Populate known sources from existing entries
            foreach (var entry in newService.LogEntries)
            {
                AddSourceIfNew(entry.Source);
            }
        }
    }

    private void OnLogEntriesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems != null)
        {
            foreach (AppLogEntry entry in e.NewItems)
            {
                AddSourceIfNew(entry.Source);
            }
        }

        // Auto-scroll to bottom
        if (AutoScrollCheckBox.IsChecked == true && LogListView.Items.Count > 0)
        {
            LogListView.ScrollIntoView(LogListView.Items[LogListView.Items.Count - 1]);
        }
    }

    private void AddSourceIfNew(string source)
    {
        if (_knownSources.Add(source))
        {
            SourceFilter.Items.Add(new ComboBoxItem { Content = source });
        }
    }

    private void SourceFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SourceFilter.SelectedItem is ComboBoxItem item)
        {
            _sourceFilter = item.Content?.ToString() ?? "All";
            RefreshFilter();
        }
    }

    private void LevelFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LevelFilter.SelectedItem is ComboBoxItem item)
        {
            _levelFilter = item.Content?.ToString() ?? "All";
            RefreshFilter();
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is IAppLogService logService)
        {
            logService.Clear();
        }
    }

    private void CopyButton_Click(object sender, RoutedEventArgs e)
    {
        var sb = new StringBuilder();
        foreach (var item in LogListView.Items)
        {
            if (item is AppLogEntry entry)
            {
                sb.AppendLine($"{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff}\t[{entry.Level}]\t[{entry.Source}]\t{entry.Message}");
            }
        }

        if (sb.Length > 0)
        {
            Clipboard.SetText(sb.ToString());
        }
    }

    private void RefreshFilter()
    {
        if (Resources["FilteredLogEntries"] is CollectionViewSource cvs)
        {
            cvs.View?.Refresh();
        }
    }

    private void FilteredLogEntries_Filter(object sender, FilterEventArgs e)
    {
        if (e.Item is not AppLogEntry entry)
        {
            e.Accepted = false;
            return;
        }

        if (_sourceFilter != "All" && entry.Source != _sourceFilter)
        {
            e.Accepted = false;
            return;
        }

        if (_levelFilter != "All")
        {
            var levelMatch = _levelFilter switch
            {
                "Debug" => entry.Level == AppLogLevel.Debug,
                "Info" => entry.Level == AppLogLevel.Info,
                "Warning" => entry.Level == AppLogLevel.Warning,
                "Error" => entry.Level == AppLogLevel.Error,
                _ => true
            };

            if (!levelMatch)
            {
                e.Accepted = false;
                return;
            }
        }

        e.Accepted = true;
    }
}
