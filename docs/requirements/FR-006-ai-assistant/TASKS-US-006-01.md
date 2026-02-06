# TASKS: US-006-01 — Start MCP Server

> **Parent US**: [US-006-01](US-006-01-start-mcp-server.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 6 | **Effort**: 2S + 2M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure and CommunityToolkit.Mvvm must be in place)
- [ ] US-008-06 (App Lifecycle - application must be initialized before MCP server can start)
- [ ] `MCPSSEServer` exists in `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` (376 lines, skeleton in place)
- [ ] `MCPToolRegistry` exists in `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` (413 lines, 7 core tools registered)
- [ ] `IGUIProxy` interface exists in `OdooAutoCAD.Core/Threading/IGUIProxy.cs` (195 lines)

## Acceptance Criteria
- [ ] AC-01: Clicking the "Start Server" button launches the MCP SSE server on the configured port (default 8084)
- [ ] AC-02: The Start button toggles to "Stop Server" once the server is running
- [ ] AC-03: The Start button is disabled while the server is already running
- [ ] AC-04: Server start verifies port availability before attempting to bind; if the port is in use, an error message is displayed: "Port {port} is already in use. Please choose a different port or stop the conflicting application."
- [ ] AC-05: All MCP tools are registered with the tool registry before the server accepts connections
- [ ] AC-06: The server gracefully handles start failure by remaining in the stopped state and logging the error
- [ ] AC-07: The status indicator changes from red (Stopped) to green (Running) upon successful start
- [ ] AC-08: The server exposes GET /sse, POST /messages, GET /health, and GET / endpoints after startup

---

## TASK-006-01-01: Implement StartServerCommand in MCPViewModel with port validation and error handling

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-01-06 |

### What to do
- Create `MCPViewModel` class extending `ObservableObject` (CommunityToolkit.Mvvm) in namespace `OdooAutoCAD.App.ViewModels`
- Add `[ObservableProperty]` fields: `bool _isServerRunning`, `string _serverStatus` (default "Stopped"), `int _serverPort` (default 8084)
- Inject `MCPSSEServer`, `MCPToolRegistry`, and `IGUIProxy` via constructor; store as `private readonly` fields
- Define `IAsyncRelayCommand StartServerCommand` using `[RelayCommand(CanExecute = nameof(CanStartServer))]`
- Implement `CanStartServer()` returning `!IsServerRunning`
- In `StartServer()`: validate port range (1024-65535, VR-006-001), check port availability using `TcpListener` or `Socket.Bind` test (VR-006-002), call `_toolRegistry.RegisterAllTools()`, then `await _mcpServer.StartAsync()`, update `IsServerRunning = true`, `ServerStatus = "Running"`
- Catch `SocketException` / `IOException` for port conflicts and display "Port {port} is already in use. Please choose a different port or stop the conflicting application."
- On any start failure, keep `IsServerRunning = false`, `ServerStatus = "Stopped"`, log the error

### How to verify
- [ ] Clicking Start launches server on configured port (AC-01)
- [ ] Port validation rejects ports outside 1024-65535 (AC-04)
- [ ] Port-in-use error displays correct message (AC-04)
- [ ] Start failure keeps server in stopped state (AC-06)

---

## TASK-006-01-02: Add Start button with toggle binding and disabled state in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Create `MCPAssistantPage.xaml` as a WPF `Page` in `OdooAutoCAD.App.Views.Pages` namespace
- Add code-behind that sets `DataContext` to `MCPViewModel` resolved from DI container
- Add a `Button` in the Server Control panel section with `Content` bound via `BoolToStartStopConverter`: shows "Start Server" when stopped, "Stop Server" when running
- Bind `Command="{Binding StartServerCommand}"` for the start action
- The button's `IsEnabled` is automatically managed by `IAsyncRelayCommand.CanExecute`
- Apply consistent styling with other pages (BackgroundBrush, PrimaryBrush, etc.)

### How to verify
- [ ] Button shows "Start Server" when server is stopped (AC-02)
- [ ] Button shows "Stop Server" when server is running (AC-02)
- [ ] Button is disabled while server is already running for start action (AC-03)

---

## TASK-006-01-03: Implement MCPSSEServer.StartAsync() with port availability check and endpoint registration

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-006-01-04 |

### What to do
- Enhance the existing `StartAsync()` method with port availability pre-check: attempt to bind a `TcpListener` to the port, release it, then proceed; if `SocketException` is thrown, propagate with descriptive message
- Add a `StatusChanged` event: `event Action<string, bool>?` that fires `(message, isRunning)` on start/stop transitions
- Ensure the `WebApplication` is configured with `builder.WebHost.UseUrls($"http://localhost:{Port}")` for explicit port binding
- Confirm all four endpoints are mapped: `GET /` (HandleRootAsync), `GET /health` (HandleHealthAsync), `GET /sse` (HandleSSEAsync), `POST /messages` (HandleMessagesAsync)
- Start the server on a background thread via `Task.Run(() => _app.RunAsync(_serverCts.Token))`
- Fire `StatusChanged?.Invoke("MCP SSE Server started on port {Port}", true)` after successful start
- Ensure idempotent behavior: if `IsRunning` is already true, return without error (VR-006-003)
- Add a `ServerStartTime` property of type `DateTime?` set to `DateTime.UtcNow` on start

### How to verify
- [ ] Server starts and exposes all four HTTP endpoints (AC-08)
- [ ] Port availability is checked before binding (AC-04)
- [ ] StatusChanged event fires with running=true on success (AC-07)
- [ ] Idempotent start returns without error when already running (AC-06)

---

## TASK-006-01-04: Register all MCP tools via MCPToolRegistry.RegisterAllTools() during server startup

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` |
| Estimate | L |
| Depends On | TASK-006-01-03 |
| Blocks | None |

### What to do
- Expand `RegisterAllTools()` to register the full set of 24 MCP tools organized by category (currently 7 tools; add remaining 17)
- Add category metadata to tool definitions: add a `Category` property to `MCPTool` model (values: "Connection", "AutoCAD Drawing", "Data Integration", "Layout", "NLP")
- Add `RequiresAutoCAD` and `RequiresOdoo` boolean properties to `MCPTool` for prerequisite tracking
- Register AutoCAD Drawing tools (8 total): create_new_drawing, draw_line, draw_circle, create_text, add_dimension, set_layer, list_layers, scan_elements -- each using `_guiProxy.ExecuteInGuiAsync()` for COM thread safety
- Register Data Integration tools (6 total): extract_autocad_parameters, sync_to_odoo, generate_boq, export_to_database, sync_drawing_to_odoo, generate_boq_from_drawing
- Register Layout tools (5 total): get_current_layout, switch_to_layout, draw_in_layout, extract_layout_parameters, export_layout_image -- all via GUI proxy
- Register NLP tool (1): process_natural_language_command
- Ensure all tool names are unique within the registry (VR-006-008)
- Add a `ToolRegistered` event for dynamic registration notifications
- Make `RegisterTool()` public to support dynamic registration per FR-006-013

### How to verify
- [ ] All 24 MCP tools are registered before server accepts connections (AC-05)
- [ ] Tool names are unique; duplicate registration throws or logs warning (AC-05)
- [ ] Each tool has Category, RequiresAutoCAD, RequiresOdoo metadata set correctly

---

## TASK-006-01-05: Initialize IGUIProxy and start DispatcherTimer for proxy polling on server start

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml.cs` |
| Estimate | M |
| Depends On | TASK-006-01-01 |
| Blocks | None |

### What to do
- In the `MCPAssistantPage` code-behind (or in `MCPViewModel` if using a service), create a `DispatcherTimer` with `Interval = TimeSpan.FromMilliseconds(100)` matching the Python pattern
- On the `Tick` event, call `_guiProxy.ProcessRequests()` to dequeue and execute pending COM operations on the STA/GUI thread
- After processing, update `GUIProxyPendingRequests` ViewModel property with `_guiProxy.PendingRequestCount`
- Start the timer when the server starts (subscribe to `MCPSSEServer.StatusChanged` or call from `StartServerCommand`)
- Call `_guiProxy.Start()` before starting the timer
- Register all default GUI proxy handlers via a setup method: switch_layout, get_current_layout, extract_parameters, extract_layout_parameters, get_autocad_status, draw_line, draw_circle, export_layout_image
- Ensure timer runs on the UI dispatcher thread for correct STA apartment

### How to verify
- [ ] DispatcherTimer polls IGUIProxy queue every 100ms (AC-01)
- [ ] AutoCAD COM operations queued by MCP tools execute on GUI thread (AC-08)
- [ ] GUI proxy handlers are registered for all 8 default actions (AC-05)

---

## TASK-006-01-06: Wire status callback to update ViewModel properties on state change

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-01-01 |
| Blocks | None |

### What to do
- Subscribe to `MCPSSEServer.StatusChanged` event in `MCPViewModel` constructor or initialization method
- On status change callback `(message, isRunning)`: use `Application.Current.Dispatcher.InvokeAsync()` to marshal to UI thread
- Update `IsServerRunning`, `ServerStatus` ("Running"/"Stopped"), `ServerStartTime` (set on start, clear on stop)
- Call `StartServerCommand.NotifyCanExecuteChanged()` and `StopServerCommand.NotifyCanExecuteChanged()` to refresh button enabled states
- Add activity log entry for the status change event
- Update the status indicator Ellipse fill via `IsServerRunning` binding (green = running, red = stopped)

### How to verify
- [ ] Status indicator changes from red to green on successful start (AC-07)
- [ ] IsServerRunning updates to true after start (AC-01)
- [ ] Button states refresh correctly after status change (AC-02, AC-03)

---

## Dependency Graph
```
TASK-006-01-01 (ViewModel + StartCommand)
       │
       ├──▶ TASK-006-01-05 (DispatcherTimer + GUIProxy)
       └──▶ TASK-006-01-06 (Status Callback Wiring)

TASK-006-01-02 (XAML Start Button)
       (independent, parallel with 01)

TASK-006-01-03 (MCPSSEServer.StartAsync Enhancement)
       │
       └──▶ TASK-006-01-04 (Full Tool Registration)

Tasks 01-03 have no cross-dependencies and can be developed in parallel.
Task 04 depends on 03 (server must exist to register tools).
Tasks 05-06 depend on 01 (ViewModel must exist for timer/callback wiring).
Task 02 is fully independent.
```
