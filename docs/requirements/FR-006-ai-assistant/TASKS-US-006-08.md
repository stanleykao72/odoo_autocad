# TASKS: US-006-08 — View Activity Logs

> **Parent US**: [US-006-08](US-006-08-view-activity-logs.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure must be in place)
- [ ] `MCPViewModel` created with base properties (from US-006-01)
- [ ] `LogEntry` model or equivalent must be defined

## Acceptance Criteria
- [ ] AC-01: The page displays a scrollable log viewer showing server activity (connections, tool calls, errors)
- [ ] AC-02: Each log entry includes a timestamp, log level (INFO/WARNING/ERROR), and message
- [ ] AC-03: The log viewer auto-scrolls to the latest entry by default
- [ ] AC-04: The user can toggle auto-scroll on/off to manually browse the log
- [ ] AC-05: A "Clear Log" button is available to reset the log viewer
- [ ] AC-06: Log entries include a source field indicating the origin (Server, Tool, SSE, GUIProxy)
- [ ] AC-07: Server start, stop, connection events, tool executions, and errors are all logged

---

## TASK-006-08-01: Create Activity Log panel XAML with scrollable ListBox, Clear Log button, and Auto-Scroll toggle

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-08-04 |

### What to do
- Add an "Activity Log" section at the bottom of `MCPAssistantPage.xaml` using a `Border` with header TextBlock
- Inside, place a `ListBox` (or `ItemsControl` in a `ScrollViewer`) with `ItemsSource="{Binding ActivityLog}"` and `x:Name="LogListBox"`
- Define a `DataTemplate` for `LogEntry` displaying: timestamp formatted as "HH:mm:ss" (monospace font), log level in brackets "[INFO]"/"[WARNING]"/"[ERROR]" with color coding (INFO=default, WARNING=orange #FF9800, ERROR=red #F44336), and the message text
- Set `Height` or `MaxHeight` on the log panel to create a scrollable area (e.g., 200-300px)
- Add a button bar below the log area with "Clear Log" button (`Command="{Binding ClearLogCommand}"`) and "Auto-Scroll" toggle button (`IsChecked="{Binding IsAutoScrollEnabled}"`)
- Apply `ScrollViewer.VerticalScrollBarVisibility="Auto"` on the ListBox

### How to verify
- [ ] Scrollable log viewer is displayed at bottom of page (AC-01)
- [ ] Log entries show timestamp, level, and message (AC-02)
- [ ] Clear Log and Auto-Scroll buttons are present (AC-05, AC-04)

---

## TASK-006-08-02: Implement LogEntry model with Timestamp, Level, Message, and Source properties

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-006-08-03 |

### What to do
- Create a `LogEntry` class (in MCPViewModel.cs or a separate Models file) with properties: `DateTime Timestamp`, `string Level` ("INFO"/"WARNING"/"ERROR"), `string Message`, `string Source` ("Server"/"Tool"/"SSE"/"GUIProxy")
- Add a computed `string FormattedTimestamp => Timestamp.ToString("HH:mm:ss")` property
- Add a computed `string FormattedEntry => $"{FormattedTimestamp} [{Level}] {Message}"` property for display
- Add a computed `Brush LevelColor` property: INFO returns default text brush, WARNING returns `Brushes.Orange`, ERROR returns `Brushes.Red`
- Add a static factory method `LogEntry.Info(string message, string source)` for convenience, and similarly `LogEntry.Warning(...)` and `LogEntry.Error(...)`
- Set `Timestamp = DateTime.Now` in each factory method

### How to verify
- [ ] LogEntry contains Timestamp, Level, Message, Source (AC-02, AC-06)
- [ ] Factory methods create entries with correct level and timestamp (AC-02)
- [ ] LevelColor returns appropriate brush for color coding (AC-02)

---

## TASK-006-08-03: Add ActivityLog ObservableCollection and log management methods to ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | TASK-006-08-02 |
| Blocks | TASK-006-08-06 |

### What to do
- Add `ObservableCollection<LogEntry> ActivityLog` property to `MCPViewModel`
- Add `[ObservableProperty] bool _isAutoScrollEnabled` (default true)
- Implement a private method `AddLog(string message, string level, string source)` that creates a `LogEntry`, adds it to `ActivityLog` via `Dispatcher.InvokeAsync` for thread safety
- Enforce a maximum capacity of 1000 entries: when `ActivityLog.Count >= 1000`, remove the oldest entry (`ActivityLog.RemoveAt(0)`) before adding the new one
- Implement `[RelayCommand] void ClearLog()` that calls `ActivityLog.Clear()`
- Implement `[RelayCommand] void ToggleAutoScroll()` that toggles `IsAutoScrollEnabled`
- Add convenience methods: `AddInfoLog(string message, string source)`, `AddWarningLog(string message, string source)`, `AddErrorLog(string message, string source)`

### How to verify
- [ ] ActivityLog collection is populated and observable (AC-01)
- [ ] ClearLog empties the collection (AC-05)
- [ ] Maximum capacity prevents unbounded growth (AC-01)
- [ ] Thread-safe additions via Dispatcher (AC-01)

---

## TASK-006-08-04: Implement auto-scroll behavior with IsAutoScrollEnabled toggle and ScrollIntoView logic

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml.cs` |
| Estimate | S |
| Depends On | TASK-006-08-01 |
| Blocks | None |

### What to do
- In the code-behind, subscribe to `ActivityLog.CollectionChanged` event (or use an attached behavior)
- On `CollectionChanged` with `NotifyCollectionChangedAction.Add`: if `IsAutoScrollEnabled` is true, call `LogListBox.ScrollIntoView(LogListBox.Items[LogListBox.Items.Count - 1])` to scroll to the last item
- Alternatively, create an `AutoScrollBehavior` attached behavior that monitors the `ItemsSource` collection and scrolls on add
- Detect manual scrolling: when the user scrolls up manually, temporarily disable auto-scroll (set `IsAutoScrollEnabled = false` in the ViewModel); re-enable when they click the Auto-Scroll toggle button
- Manual scroll detection can use `ScrollViewer.ScrollChanged` event: if `e.VerticalOffset < e.ExtentHeight - e.ViewportHeight`, user has scrolled up

### How to verify
- [ ] Log viewer auto-scrolls to latest entry by default (AC-03)
- [ ] Auto-scroll can be toggled on/off (AC-04)
- [ ] Manual scrolling temporarily disables auto-scroll (AC-04)

---

## TASK-006-08-05: Implement ClearLogCommand and ToggleAutoScrollCommand relay commands

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Define `IRelayCommand ClearLogCommand` using `[RelayCommand]` attribute on `ClearLog()` method
- Define `IRelayCommand ToggleAutoScrollCommand` using `[RelayCommand]` attribute on `ToggleAutoScroll()` method
- `ClearLog()` implementation: call `ActivityLog.Clear()`, then add a single log entry "Log cleared" at INFO level, Source="Server"
- `ToggleAutoScroll()` implementation: flip `IsAutoScrollEnabled = !IsAutoScrollEnabled`
- Both commands should always be executable (no `CanExecute` restrictions)
- Optionally add a confirmation dialog before clearing if the log has more than 100 entries

### How to verify
- [ ] Clear Log button resets the log viewer (AC-05)
- [ ] Toggle Auto-Scroll button changes auto-scroll behavior (AC-04)

---

## TASK-006-08-06: Wire MCPSSEServer events to the ActivityLog collection

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | TASK-006-08-03 |
| Blocks | None |

### What to do
- In `MCPViewModel`, subscribe to `MCPSSEServer.StatusChanged`: log "MCP SSE Server started on port {port}" or "MCP SSE Server stopped" with Source="Server"
- Subscribe to `MCPSSEServer.ClientConnected`: log "SSE connection established: {connectionId}" with Source="SSE"
- Subscribe to `MCPSSEServer.ClientDisconnected`: log "SSE connection closed: {connectionId}" with Source="SSE"
- Subscribe to `MCPToolRegistry` execution events (or add hooks): log "Executing tool: {toolName}" with Source="Tool" on start, "Tool completed: {toolName}" or "Tool completed ({duration}s): {toolName}" with Source="Tool" on success, "Tool failed: {toolName} - {error}" at ERROR level with Source="Tool" on failure
- Subscribe to `IGUIProxy.RequestCompleted`: log "GUI Proxy action completed: {action}" with Source="GUIProxy"
- Subscribe to `IGUIProxy.ErrorOccurred`: log "GUI Proxy error: {error}" at ERROR level with Source="GUIProxy"
- All event handlers must marshal to UI thread via `Dispatcher.InvokeAsync` before modifying the ObservableCollection

### How to verify
- [ ] Server start/stop events are logged (AC-07)
- [ ] Connection events are logged with connection IDs (AC-07)
- [ ] Tool execution events are logged with names and durations (AC-07)
- [ ] Errors are logged at ERROR level (AC-07)
- [ ] Source field correctly indicates origin (AC-06)

---

## Dependency Graph
```
TASK-006-08-01 (XAML Log Panel)
       │
       └──▶ TASK-006-08-04 (Auto-Scroll Behavior)

TASK-006-08-02 (LogEntry Model)
       │
       └──▶ TASK-006-08-03 (ActivityLog Collection)
                   │
                   └──▶ TASK-006-08-06 (Event Wiring)

TASK-006-08-05 (Commands)
       (independent)

Tasks 01, 02, and 05 are independent and can be developed in parallel.
Task 04 depends on 01 (XAML must exist for scroll behavior).
Task 03 depends on 02 (LogEntry model must exist for collection).
Task 06 depends on 03 (collection must exist to add entries).
```
