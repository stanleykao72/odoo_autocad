# TASKS: US-006-10 — Restart Server

> **Parent US**: [US-006-10](US-006-10-restart-server.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 5 | **Effort**: 3S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-006-01 (Start MCP Server - StartAsync must be implemented)
- [ ] US-006-02 (Stop MCP Server - StopAsync must be implemented)
- [ ] `MCPSSEServer` exists with both `StartAsync()` and `StopAsync()` methods
- [ ] `MCPViewModel` created with `IsServerRunning` property and start/stop commands

## Acceptance Criteria
- [ ] AC-01: A "Restart" button is available in the server control panel
- [ ] AC-02: Clicking Restart stops the server, waits at least 1 second, then starts it again
- [ ] AC-03: The server status transitions through "Stopping..." -> "Stopped" -> "Starting..." -> "Running" during restart
- [ ] AC-04: All active SSE connections are closed during the stop phase before the server restarts
- [ ] AC-05: The Restart button is disabled while the server is already stopped (restart only applies when running)
- [ ] AC-06: If the stop phase succeeds but the start phase fails, the server remains in the stopped state with an error message
- [ ] AC-07: The restart operation is logged in the activity log with start and completion timestamps

---

## TASK-006-10-01: Implement RestartServerCommand in MCPViewModel that chains StopAsync, delay, and StartAsync

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-10-04, TASK-006-10-05 |

### What to do
- Define `IAsyncRelayCommand RestartServerCommand` using `[RelayCommand(CanExecute = nameof(CanRestartServer))]`
- Implement `CanRestartServer()` returning `IsServerRunning` (restart only available when running, AC-05)
- In `RestartServer()`: set `ServerStatus = "Stopping..."`, disable all server control buttons (Start, Stop, Restart, Test Connection) by setting a `_isRestarting` flag
- Call `await _mcpServer.StopAsync()`, set `ServerStatus = "Stopped"`
- Wait: `await Task.Delay(1000)` (minimum 1-second delay per VR-006-007)
- Set `ServerStatus = "Starting..."`, call `await _mcpServer.StartAsync()`
- On success: set `ServerStatus = "Running"`, `IsServerRunning = true`
- On start failure after successful stop: set `ServerStatus = "Stopped"`, `IsServerRunning = false`, display error message
- Clear `_isRestarting` flag and refresh all `CanExecute` states
- Catch and log any exceptions during the restart sequence

### How to verify
- [ ] Restart chains stop, delay, and start (AC-02)
- [ ] Status transitions through intermediate states (AC-03)
- [ ] Start failure after stop leaves server stopped with error (AC-06)

---

## TASK-006-10-02: Add Restart button to Server Control panel XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `Button` with `Content="Restart"` in the Server Control panel button bar, positioned between the Stop and Test Connection buttons
- Bind `Command="{Binding RestartServerCommand}"` -- `IsEnabled` is automatically managed by `CanExecute`
- Apply consistent button styling with other buttons in the panel
- Add a busy indicator (small `ProgressRing` or animated dots) visible when `RestartServerCommand.IsRunning` is true, indicating the restart is in progress
- Add `ToolTip="Stop and restart the MCP SSE server"` for user guidance

### How to verify
- [ ] Restart button is present in server control panel (AC-01)
- [ ] Button is disabled when server is stopped (AC-05)
- [ ] Busy indicator shows during restart operation (AC-01)

---

## TASK-006-10-03: Implement MCPSSEServer.RestartAsync() with 1-second delay between stop and start

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add a public `async Task RestartAsync()` method to `MCPSSEServer`
- Implementation: `await StopAsync()`, `await Task.Delay(1000)`, `await StartAsync()`
- Fire `StatusChanged?.Invoke("MCP SSE Server restarting...", false)` before stop
- After successful restart, fire `StatusChanged?.Invoke("MCP SSE Server restarted on port {Port}", true)`
- If `StartAsync` throws after successful `StopAsync`, catch the exception, fire `StatusChanged?.Invoke("MCP SSE Server restart failed: {error}", false)`, and rethrow for the caller to handle
- The 1-second delay ensures the OS fully releases the TCP port (VR-006-007)
- All active SSE connections are closed during the `StopAsync` phase (delegates to existing stop logic)
- If the server is not running when RestartAsync is called, throw `InvalidOperationException("Server is not running")`

### How to verify
- [ ] Restart stops server, waits 1 second, starts again (AC-02)
- [ ] All SSE connections are closed during stop phase (AC-04)
- [ ] Port conflict after restart is handled (VR-006-007)

---

## TASK-006-10-04: Update ServerStatus property through intermediate states during restart

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-10-01 |
| Blocks | None |

### What to do
- In the `RestartServer()` command method, update `ServerStatus` at each stage:
  1. Before stop: `ServerStatus = "Stopping..."`
  2. After stop completes: `ServerStatus = "Stopped"` (briefly visible)
  3. Before start: `ServerStatus = "Starting..."`
  4. After start completes: `ServerStatus = "Running"`
  5. On failure: `ServerStatus = "Stopped"` with error context
- Ensure each status update triggers `OnPropertyChanged` for XAML binding refresh
- The status text is bound to a TextBlock in the Server Control panel that shows the current state
- During the "Stopping..." and "Starting..." phases, the status indicator ellipse should show an intermediate color (yellow/amber) or blink -- implement with a `DataTrigger` on `ServerStatus` values

### How to verify
- [ ] Status transitions through all four states (AC-03)
- [ ] Each state is visible in the UI (AC-03)
- [ ] Failure state shows "Stopped" (AC-06)

---

## TASK-006-10-05: Log restart operation start and completion in activity log

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-10-01 |
| Blocks | None |

### What to do
- At the beginning of `RestartServer()`, add log entry: "Restarting MCP SSE Server..." at INFO level, Source="Server", with `DateTime.Now` timestamp
- After successful restart, add log entry: "MCP SSE Server restarted successfully on port {ServerPort}" at INFO level, Source="Server", with `DateTime.Now` timestamp
- If restart fails, add log entry: "MCP SSE Server restart failed: {error}" at ERROR level, Source="Server"
- Calculate and log the total restart duration: capture start time before stop, end time after start, log "Restart completed in {duration}ms"
- The intermediate stop and start events are already logged by the StatusChanged handler (from US-006-08)
- Log entries should appear in chronological order in the activity log

### How to verify
- [ ] Restart start is logged with timestamp (AC-07)
- [ ] Restart completion is logged with timestamp (AC-07)
- [ ] Restart failure is logged with error details (AC-07)

---

## Dependency Graph
```
TASK-006-10-01 (RestartServerCommand)
       │
       ├──▶ TASK-006-10-04 (Intermediate Status States)
       └──▶ TASK-006-10-05 (Restart Logging)

TASK-006-10-02 (XAML Restart Button)
       (independent)

TASK-006-10-03 (MCPSSEServer.RestartAsync)
       (independent)

Tasks 01, 02, and 03 are independent and can be developed in parallel.
Tasks 04 and 05 depend on 01 (RestartServerCommand must exist).
```
