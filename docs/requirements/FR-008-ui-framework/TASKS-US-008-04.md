# TASKS: US-008-04 — System Log Panel

> **Parent US**: [US-008-04](US-008-04-system-log-panel.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 4S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-008-01 completed (main window layout must include the content area where the status bar and log panel reside)

## Acceptance Criteria
- [ ] AC-01: A status bar at the bottom of the content area displays a `StatusMessage` text, the current date/time, and the application version.
- [ ] AC-02: A dedicated system log panel is available (as a collapsible bottom panel or a separate toggle overlay) to display operational log messages.
- [ ] AC-03: The system log panel auto-scrolls to the latest entry when new messages are added.
- [ ] AC-04: Log entries include timestamps in the format `[YYYY-MM-DD HH:mm:ss]` and category prefixes (e.g., "[SSE GUI]", "[GUI Proxy]", "[Odoo]").
- [ ] AC-05: The status bar displays three columns: status message (left, fills remaining space), date/time (center-right), and version string (far right).
- [ ] AC-06: The status bar uses `SurfaceBrush` background with a top border separator.

---

## TASK-008-04-01: Define status bar XAML with three-column layout

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the content area Grid Row 2, define a `Border` with `Background="{StaticResource SurfaceBrush}"`, `Padding="15,8"`, and `BorderBrush="{StaticResource BorderBrush}"` with `BorderThickness="0,1,0,0"` (top border)
- Add an inner `Grid` with three `ColumnDefinitions`: `*` (status message fills), Auto (date/time), Auto (version)
- Add `TextBlock` for `StatusMessage` in column 0 bound to `{Binding StatusMessage}` with `TextSecondaryBrush` foreground
- Add `TextBlock` for `CurrentDateTime` in column 1 bound to `{Binding CurrentDateTime}` with `Margin="20,0"` and `TextSecondaryBrush`
- Add `TextBlock` for version string "v6.0.0" in column 2 with `TextSecondaryBrush`
- Verify the skeleton already defines this layout; confirm correctness

### How to verify
- [ ] Status bar shows three columns: message, date/time, version (AC-05)
- [ ] Status bar uses SurfaceBrush background with top border (AC-06)
- [ ] StatusMessage text is bound to MainViewModel property (AC-01)

---

## TASK-008-04-02: Add StatusMessage, CurrentDateTime, and version observable properties to MainViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-008-04-06 |

### What to do
- Verify `[ObservableProperty] private string _statusMessage = "Ready";` exists in `MainViewModel`
- Verify `[ObservableProperty] private string _currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");` exists
- Both should already be in the skeleton; validate they are present and raise `PropertyChanged`
- Optionally add a constant or `[ObservableProperty]` for the version string if dynamic display is needed; otherwise the static "v6.0.0" in XAML suffices

### How to verify
- [ ] `StatusMessage` property exists with default "Ready" (AC-01)
- [ ] `CurrentDateTime` property exists with formatted date string (AC-01)
- [ ] Properties raise PropertyChanged for real-time binding updates (AC-01)

---

## TASK-008-04-03: Implement collapsible log panel XAML with ScrollViewer

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | M |
| Depends On | TASK-008-04-01 |
| Blocks | TASK-008-04-05 |

### What to do
- Add a fourth row to the content area Grid: `<RowDefinition Height="Auto"/>` below the status bar for the log panel
- Use an `Expander` or a `Border` with a toggle `Button` to make the panel collapsible; set `IsExpanded="False"` by default
- Inside the collapsible region, add a `Border` with dark background (black or `#1E1E1E`) and fixed `MaxHeight="200"`
- Add a `ScrollViewer` named `LogScrollViewer` with `VerticalScrollBarVisibility="Auto"` wrapping an `ItemsControl` or `TextBlock`
- Bind the `ItemsControl.ItemsSource` to `{Binding LogEntries}` (an `ObservableCollection<string>` on MainViewModel)
- Style log entries with monospaced font (Consolas), white/light-gray text on dark background
- Add a header label "System Log" or integrate into the Expander header

### How to verify
- [ ] System log panel is available as a collapsible section (AC-02)
- [ ] Log panel displays operational messages (AC-02)
- [ ] Panel is hidden by default and toggleable (AC-02)

---

## TASK-008-04-04: Create Serilog sink or log collection for UI panel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | M |
| Depends On | TASK-008-04-03 |
| Blocks | TASK-008-04-05 |

### What to do
- Add `ObservableCollection<string> LogEntries` property to `MainViewModel`
- Implement an `AddLogEntry(string category, string message)` method that creates formatted entries: `$"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{category}] {message}"`
- Optionally implement a custom Serilog sink (`InMemorySink` or `ObservableCollectionSink`) that writes to the `LogEntries` collection via `Dispatcher.Invoke` for thread safety
- Register the sink in `App.OnStartup` Serilog configuration: `.WriteTo.Sink(uiSink)`
- Alternatively, expose a static `Action<string>` callback that Serilog writes to, and have `MainViewModel` subscribe
- Ensure category prefixes match Python conventions: "[SSE GUI]", "[GUI Proxy]", "[Odoo]", "[AutoCAD]"

### How to verify
- [ ] Log entries appear in the UI panel with timestamps (AC-04)
- [ ] Category prefixes are included in log entries (AC-04)
- [ ] Timestamp format is `[YYYY-MM-DD HH:mm:ss]` (AC-04)

---

## TASK-008-04-05: Implement auto-scroll behavior on log panel ScrollViewer

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml.cs` |
| Estimate | S |
| Depends On | TASK-008-04-03, TASK-008-04-04 |
| Blocks | None |

### What to do
- Subscribe to the `LogEntries.CollectionChanged` event in `MainWindow` code-behind (or use an attached behavior)
- On each `CollectionChanged` with `NotifyCollectionChangedAction.Add`, call `LogScrollViewer.ScrollToEnd()` to auto-scroll
- Alternatively, use a WPF attached behavior or `Behavior<ScrollViewer>` from `Microsoft.Xaml.Behaviors` for MVVM purity
- Ensure the scroll occurs on the dispatcher thread using `Dispatcher.BeginInvoke` if needed

### How to verify
- [ ] Log panel auto-scrolls to the latest entry when new messages are added (AC-03)
- [ ] User can still manually scroll up to view older entries (AC-03)

---

## TASK-008-04-06: Add DispatcherTimer to update CurrentDateTime every second

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` |
| Estimate | S |
| Depends On | TASK-008-04-02 |
| Blocks | None |

### What to do
- Verify a `DispatcherTimer` with 1-second interval exists in `MainViewModel` (`_statusTimer`)
- In the tick handler `OnStatusTimerTick`, update `CurrentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")`
- The skeleton already implements this; confirm the timer starts in the constructor and updates the property
- Ensure the timer is stopped when `MainViewModel` is disposed or the application shuts down

### How to verify
- [ ] Date/time in the status bar updates every second (AC-01)
- [ ] Format matches `yyyy-MM-dd HH:mm:ss` (AC-01)

---

## Dependency Graph

```
TASK-008-04-01 (Status bar XAML) [independent]
    └── TASK-008-04-03 (Log panel XAML)
        └── TASK-008-04-05 (Auto-scroll)

TASK-008-04-02 (ViewModel properties)
    └── TASK-008-04-06 (DateTime timer)

TASK-008-04-04 (Serilog sink)
    └── TASK-008-04-05 (Auto-scroll)
```
