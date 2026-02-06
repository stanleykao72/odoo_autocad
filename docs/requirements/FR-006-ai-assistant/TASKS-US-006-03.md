# TASKS: US-006-03 — View Server Status

> **Parent US**: [US-006-03](US-006-03-view-server-status.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure must be in place)
- [ ] `MCPSSEServer` exists with `IsRunning`, `Port`, `ActiveConnections` properties
- [ ] `MCPViewModel` created in US-006-01 with base properties

## Acceptance Criteria
- [ ] AC-01: The page displays a visual status indicator (green ellipse for Running, red ellipse for Stopped)
- [ ] AC-02: The current server port number is displayed in the status panel
- [ ] AC-03: The number of active SSE connections is displayed and updated in real-time
- [ ] AC-04: The number of registered MCP tools is displayed in the status panel
- [ ] AC-05: The MCP session initialized state (whether a client has completed the handshake) is displayed
- [ ] AC-06: Status updates are pushed to the UI in real-time via a callback mechanism when server state changes
- [ ] AC-07: Server uptime is displayed when the server is running (formatted as Xh Ym Zs)

---

## TASK-006-03-01: Create Server Control panel XAML layout with status indicator, port, uptime, connections, and tool count

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-03-02 |

### What to do
- Add a `Border` with `CornerRadius="8"` in the top section of `MCPAssistantPage.xaml` for the Server Control panel
- Inside, use a `Grid` with rows/columns to lay out: status indicator `Ellipse` (Width=16, Height=16) with `Fill` bound to color, "Status:" label + status text TextBlock, "Port:" label + port value TextBlock, "Uptime:" label + uptime TextBlock, "Connections:" label + connection count TextBlock, "Protocol:" label + "MCP 2024-11-05" TextBlock, "Tools:" label + tool count TextBlock + "registered" suffix
- Add "Session Initialized:" label with a boolean indicator (checkmark or text "Yes"/"No") bound to `IsSessionInitialized`
- Place the Start/Stop, Restart, and Test Connection buttons in a `StackPanel Orientation="Horizontal"` below the status info
- Use `{Binding ServerPort}`, `{Binding ServerUptime}`, `{Binding ActiveConnectionCount}`, `{Binding RegisteredToolCount}`, `{Binding IsSessionInitialized}` for data bindings
- Follow the wireframe layout from FR-006 Section 5

### How to verify
- [ ] Status indicator ellipse renders with correct color binding (AC-01)
- [ ] Port number is displayed in the panel (AC-02)
- [ ] Connection count, tool count, session state, and uptime labels are present and bound (AC-03, AC-04, AC-05, AC-07)

---

## TASK-006-03-02: Implement BoolToColorConverter for mapping IsServerRunning to green/red Ellipse fill

| Field | Value |
|-------|-------|
| Target | `Converters/BoolToColorConverter.cs` |
| Estimate | S |
| Depends On | TASK-006-03-01 |
| Blocks | None |

### What to do
- Create `BoolToColorConverter` implementing `IValueConverter` in `OdooAutoCAD.App.Converters` namespace (if not already created in US-001-01)
- `Convert`: return `new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50))` (green #4CAF50) when `true`; return `new SolidColorBrush(Color.FromRgb(0xF4, 0x43, 0x36))` (red #F44336) when `false`
- `ConvertBack`: throw `NotSupportedException`
- Register in `MCPAssistantPage.xaml` resources: `<converters:BoolToColorConverter x:Key="BoolToColorConverter"/>`
- Bind the status Ellipse: `Fill="{Binding IsServerRunning, Converter={StaticResource BoolToColorConverter}}"`
- If the converter already exists from FR-001, reuse it as a shared resource

### How to verify
- [ ] Green ellipse when server is running; red ellipse when stopped (AC-01)

---

## TASK-006-03-03: Add ViewModel properties for server status display

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-03-04, TASK-006-03-05 |

### What to do
- Add `[ObservableProperty]` fields to `MCPViewModel` (if not already present from US-006-01): `int _activeConnectionCount` (default 0), `int _registeredToolCount` (default 0), `bool _isSessionInitialized` (default false), `string _serverUptime` (default "--"), `DateTime? _serverStartTime`
- Add `string _protocolVersion` property (constant "2024-11-05")
- Add `string _sseEndpointUrl` computed from `ServerPort`
- Override `OnPropertyChanged` for `ServerPort` to recalculate `SSEEndpointUrl`
- Initialize `RegisteredToolCount` from `_toolRegistry.GetTools().Count` in the constructor
- Ensure all properties fire `PropertyChanged` notifications via CommunityToolkit `[ObservableProperty]`

### How to verify
- [ ] All status properties are observable and bindable (AC-02, AC-03, AC-04, AC-05)
- [ ] RegisteredToolCount reflects actual tool registry count (AC-04)

---

## TASK-006-03-04: Implement status callback handler that updates ViewModel on server state changes

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-03-03 |
| Blocks | None |

### What to do
- Subscribe to `MCPSSEServer.StatusChanged` event in `MCPViewModel` initialization
- Create a private method `OnServerStatusChanged(string message, bool isRunning)` that marshals to UI thread via `Dispatcher.InvokeAsync`
- Update `IsServerRunning`, `ServerStatus`, `ActiveConnectionCount` (from `_mcpServer.ActiveConnections`), `IsSessionInitialized` (from server state)
- When `isRunning` transitions to true: set `ServerStartTime = DateTime.UtcNow`, start uptime timer
- When `isRunning` transitions to false: clear `ServerStartTime`, set `ServerUptime = "--"`, stop uptime timer, set `ActiveConnectionCount = 0`
- Refresh command `CanExecute` states for all server control commands
- Add an activity log entry for the status change

### How to verify
- [ ] Status updates in real-time when server state changes (AC-06)
- [ ] Connection count and session state update on transitions (AC-03, AC-05)

---

## TASK-006-03-05: Implement uptime timer that updates ServerUptime every second while running

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-03-03 |
| Blocks | None |

### What to do
- Create a `DispatcherTimer _uptimeTimer` with `Interval = TimeSpan.FromSeconds(1)` in `MCPViewModel`
- On `Tick`: if `ServerStartTime.HasValue`, calculate `elapsed = DateTime.UtcNow - ServerStartTime.Value`, format as `"{hours}h {minutes}m {seconds}s"` (e.g., "1h 23m 45s"), assign to `ServerUptime`
- Start the timer when `IsServerRunning` becomes true (call from `OnServerStatusChanged`)
- Stop the timer when `IsServerRunning` becomes false; set `ServerUptime = "--"`
- Handle edge case where elapsed is zero or negative (display "0h 0m 0s")
- Implement `IDisposable` pattern: stop the timer in `Dispose()` to prevent memory leaks

### How to verify
- [ ] Uptime displays formatted time while server is running (AC-07)
- [ ] Uptime shows "--" when server is stopped (AC-07)
- [ ] Uptime updates every second without manual refresh (AC-06)

---

## TASK-006-03-06: Expose ActiveConnections count and IsInitialized from MCPSSEServer for ViewModel binding

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add a public `bool IsInitialized => _isInitialized` property to `MCPSSEServer`
- Add a public `DateTime? ServerStartTime` property, set in `StartAsync()` and cleared in `StopAsync()`
- Add an `event Action<int>? ConnectionCountChanged` that fires when connections are added or removed
- In `HandleSSEAsync`, after adding a connection to `_connections`, fire `ConnectionCountChanged?.Invoke(ActiveConnections)`
- In the `finally` block of `HandleSSEAsync` (connection removal), fire `ConnectionCountChanged?.Invoke(ActiveConnections)`
- In `StopAsync`, after clearing connections, fire `ConnectionCountChanged?.Invoke(0)`
- Add a `MCPClientInfo? ClientInfo => _clientInfo` property to expose client info for the status panel
- These properties allow `MCPViewModel` to poll or subscribe for real-time status display

### How to verify
- [ ] ActiveConnections accurately reflects ConcurrentDictionary count (AC-03)
- [ ] IsInitialized reflects handshake completion state (AC-05)
- [ ] ConnectionCountChanged fires on connect/disconnect events (AC-03, AC-06)

---

## Dependency Graph
```
TASK-006-03-01 (XAML Server Control Panel)
       │
       └──▶ TASK-006-03-02 (BoolToColorConverter)

TASK-006-03-03 (ViewModel Status Properties)
       │
       ├──▶ TASK-006-03-04 (Status Callback Handler)
       └──▶ TASK-006-03-05 (Uptime Timer)

TASK-006-03-06 (MCPSSEServer Properties/Events)
       (independent - can be developed in parallel)

Tasks 01, 03, and 06 have no cross-dependencies.
Task 02 depends on 01 (XAML must exist to register converter).
Tasks 04-05 depend on 03 (ViewModel properties must exist).
```
