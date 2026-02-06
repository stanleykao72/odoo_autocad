# TASKS: US-006-02 — Stop MCP Server

> **Parent US**: [US-006-02](US-006-02-stop-mcp-server.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 5 | **Effort**: 3S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-006-01 (Start MCP Server - server must be running to stop it; StartServerCommand and MCPViewModel must exist)
- [ ] `MCPSSEServer.StopAsync()` skeleton exists in `OdooAutoCAD.MCP/Server/MCPSSEServer.cs`
- [ ] `IGUIProxy` interface exists in `OdooAutoCAD.Core/Threading/IGUIProxy.cs`

## Acceptance Criteria
- [ ] AC-01: Clicking the "Stop Server" button shuts down the MCP SSE server and releases the port
- [ ] AC-02: All active SSE connections are gracefully closed before the server shuts down
- [ ] AC-03: The Stop button toggles back to "Start Server" once the server has stopped
- [ ] AC-04: The Stop button is disabled while the server is already stopped
- [ ] AC-05: The status indicator changes from green (Running) to red (Stopped) upon successful stop
- [ ] AC-06: The DispatcherTimer for GUI Proxy polling is stopped when the server stops
- [ ] AC-07: If the server is already stopped, the stop operation is idempotent and returns without error

---

## TASK-006-02-01: Implement StopServerCommand in MCPViewModel with graceful shutdown logic

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-02-05 |

### What to do
- Define `IAsyncRelayCommand StopServerCommand` using `[RelayCommand(CanExecute = nameof(CanStopServer))]`
- Implement `CanStopServer()` returning `IsServerRunning`
- In `StopServer()`: set `ServerStatus = "Stopping..."`, call `await _mcpServer.StopAsync()`, update `IsServerRunning = false`, `ServerStatus = "Stopped"`, clear `ServerStartTime`
- Call `NotifyCanExecuteChanged()` on both `StartServerCommand` and `StopServerCommand` after state change
- Catch exceptions during stop and log them; ensure server transitions to stopped state even on error
- Handle idempotent stop: if `!IsServerRunning`, return immediately without error (VR-006-004)

### How to verify
- [ ] Clicking Stop shuts down the server and releases the port (AC-01)
- [ ] Stop is idempotent when server is already stopped (AC-07)
- [ ] ServerStatus transitions through "Stopping..." to "Stopped" (AC-05)

---

## TASK-006-02-02: Enhance MCPSSEServer.StopAsync() to close all SSE connections and release port

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Enhance the existing `StopAsync()` to log the number of active connections being closed before clearing
- Send a "closing" SSE event to each connected client before cancelling their `CancellationTokenSource`
- Iterate `_connections.Values`, send close event, then cancel; use `Task.WhenAll` with per-connection timeout (2s) to avoid hanging on unresponsive clients
- Clear the `ConcurrentDictionary<string, SSEConnection>` after all connections are closed
- Cancel `_serverCts`, call `await _app.StopAsync()`, dispose the `WebApplication`, set `_app = null` and `_serverTask = null`
- Reset `_isInitialized = false` and `_clientInfo = null`
- Fire `StatusChanged?.Invoke("MCP SSE Server stopped", false)` after successful stop
- Clear `ServerStartTime` (set to null)
- Ensure idempotent behavior: if `!IsRunning`, return without error (VR-006-004)

### How to verify
- [ ] All active SSE connections are gracefully closed before shutdown (AC-02)
- [ ] Port is released after stop (AC-01)
- [ ] StatusChanged fires with isRunning=false (AC-05)
- [ ] Idempotent stop returns without error (AC-07)

---

## TASK-006-02-03: Stop DispatcherTimer and clean up IGUIProxy resources on server stop

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Subscribe to `MCPSSEServer.StatusChanged` or `MCPViewModel.PropertyChanged` for `IsServerRunning` changes
- When server stops (`isRunning == false`): call `_proxyTimer.Stop()` to halt the 100ms DispatcherTimer
- Call `_guiProxy.ClearPendingRequests()` to discard any queued but unprocessed requests
- Call `_guiProxy.Stop()` to cleanly shut down the proxy service
- Update `IsGUIProxyRunning` and `GUIProxyPendingRequests` ViewModel properties to reflect stopped state
- Ensure the timer is not started again until the next server start

### How to verify
- [ ] DispatcherTimer is stopped when server stops (AC-06)
- [ ] GUI Proxy pending requests are cleared (AC-06)
- [ ] No COM operations are attempted after server stops

---

## TASK-006-02-04: Update XAML button binding to show "Start Server" when stopped and disable Stop while stopped

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Create a `BoolToStartStopConverter` in `OdooAutoCAD.App.Converters` that returns "Stop Server" when `true` and "Start Server" when `false`
- Register the converter in `MCPAssistantPage.xaml` resources
- Bind the toggle button `Content` to `{Binding IsServerRunning, Converter={StaticResource BoolToStartStopConverter}}`
- Bind the toggle button `Command` conditionally: when running, bind to `StopServerCommand`; when stopped, bind to `StartServerCommand` (or use a single `ToggleServerCommand` that delegates)
- Alternative approach: use two separate buttons with `Visibility` toggled by `IsServerRunning` binding
- Ensure the Stop button is disabled when the server is already stopped via `CanExecute` binding (AC-04)

### How to verify
- [ ] Button shows "Start Server" when stopped (AC-03)
- [ ] Button shows "Stop Server" when running (AC-03)
- [ ] Stop button is disabled while server is stopped (AC-04)

---

## TASK-006-02-05: Log server shutdown event and connection closures in activity log

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-02-01 |
| Blocks | None |

### What to do
- In the `StopServer()` method, before calling `StopAsync()`, add a log entry: "Stopping MCP SSE Server..." at INFO level with Source="Server"
- After successful stop, add log entry: "MCP SSE Server stopped. {N} connections closed." at INFO level with Source="Server"
- The connection count should be captured before the stop (from `_mcpServer.ActiveConnections`) and included in the log message
- If stop fails with an exception, add log entry at ERROR level: "MCP Server stop failed: {error}" with Source="Server"
- Use `Application.Current.Dispatcher.InvokeAsync()` to marshal log additions to the UI thread for `ObservableCollection` thread safety

### How to verify
- [ ] Server shutdown is logged in the activity log (AC-01)
- [ ] Number of closed connections is included in the log message (AC-02)
- [ ] Status indicator changes from green to red (AC-05)

---

## Dependency Graph
```
TASK-006-02-01 (StopServerCommand)
       │
       └──▶ TASK-006-02-05 (Shutdown Logging)

TASK-006-02-02 (MCPSSEServer.StopAsync Enhancement)
       (independent)

TASK-006-02-03 (DispatcherTimer Cleanup)
       (independent)

TASK-006-02-04 (XAML Button Binding)
       (independent)

Tasks 01-04 have no cross-dependencies and can be developed in parallel.
Task 05 depends on 01 (StopServerCommand must exist to add logging).
```
