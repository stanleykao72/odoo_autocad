# TASKS: US-006-11 — View Tool Execution

> **Parent US**: [US-006-11](US-006-11-view-tool-execution.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 5 | **Effort**: 3S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-006-01 (Start MCP Server - tools can only execute when the server is running)
- [ ] US-006-08 (View Activity Logs - execution results are displayed in the activity log)
- [ ] `MCPToolRegistry.ExecuteToolAsync()` exists with tool execution capability
- [ ] `ActivityLog` ObservableCollection exists in `MCPViewModel`

## Acceptance Criteria
- [ ] AC-01: When an MCP tool is invoked by an AI assistant, the tool name appears in the activity log with an "Executing" status
- [ ] AC-02: When the tool completes, the log entry updates to show the execution duration and success/failure status
- [ ] AC-03: Tool execution results include the tool name, execution time in seconds, and outcome (success or error description)
- [ ] AC-04: Failed tool executions are logged at ERROR level with the failure reason
- [ ] AC-05: GUI Proxy operations (AutoCAD COM calls) show the queued action name and completion status
- [ ] AC-06: The activity log updates in real-time as tools execute, without requiring manual refresh

---

## TASK-006-11-01: Add execution logging hooks in MCPToolRegistry.ExecuteToolAsync()

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-11-02 |

### What to do
- Add events to `MCPToolRegistry` for tool execution lifecycle: `event Action<string>? ToolExecutionStarted` (fires with tool name), `event Action<string, TimeSpan, bool>? ToolExecutionCompleted` (fires with tool name, duration, success), `event Action<string, string>? ToolExecutionFailed` (fires with tool name, error message)
- In `ExecuteToolAsync()`, before calling `executor(arguments)`: fire `ToolExecutionStarted?.Invoke(toolName)`, start a `System.Diagnostics.Stopwatch`
- After successful execution: stop the Stopwatch, fire `ToolExecutionCompleted?.Invoke(toolName, stopwatch.Elapsed, true)`
- In the `catch` block on failure: stop the Stopwatch, fire `ToolExecutionFailed?.Invoke(toolName, ex.Message)` and `ToolExecutionCompleted?.Invoke(toolName, stopwatch.Elapsed, false)`
- Keep the existing `_logger` calls for internal logging alongside the new events
- Ensure events fire regardless of whether subscribers exist (use null-conditional invocation)

### How to verify
- [ ] ToolExecutionStarted fires before tool execution (AC-01)
- [ ] ToolExecutionCompleted fires with duration and success status (AC-02, AC-03)
- [ ] ToolExecutionFailed fires with error message on failure (AC-04)

---

## TASK-006-11-02: Implement tool execution event handler in MCPViewModel to append LogEntry items

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-11-01 |
| Blocks | None |

### What to do
- Subscribe to `MCPToolRegistry.ToolExecutionStarted` in MCPViewModel initialization
- On start event: call `AddInfoLog($"Executing tool: {toolName}", "Tool")` to add a log entry
- Subscribe to `MCPToolRegistry.ToolExecutionCompleted`
- On completion: if success, call `AddInfoLog($"Tool completed ({duration.TotalSeconds:F1}s): {toolName}", "Tool")`; if duration < 0.1s, omit the duration for brevity: `AddInfoLog($"Tool completed: {toolName}", "Tool")`
- Subscribe to `MCPToolRegistry.ToolExecutionFailed`
- On failure: call `AddErrorLog($"Tool failed: {toolName} - {errorMessage}", "Tool")`
- All log additions must be marshalled to the UI thread via `Dispatcher.InvokeAsync` for ObservableCollection thread safety
- The log entries update in real-time as events arrive (AC-06)

### How to verify
- [ ] "Executing tool: {name}" appears in log on tool start (AC-01)
- [ ] "Tool completed ({duration}s): {name}" appears on success (AC-02, AC-03)
- [ ] "Tool failed: {name} - {error}" appears at ERROR level on failure (AC-04)
- [ ] Log updates without manual refresh (AC-06)

---

## TASK-006-11-03: Add Stopwatch-based duration measurement in ExecuteToolAsync

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add `using System.Diagnostics;` to the file imports
- In `ExecuteToolAsync()`, create `var stopwatch = Stopwatch.StartNew();` before calling the executor
- After execution (success or failure), call `stopwatch.Stop()`
- Pass `stopwatch.Elapsed` to the `ToolExecutionCompleted` event
- Optionally add the execution duration to the `MCPToolCallResult` metadata (add a `Duration` property to the result model if desired for downstream consumers)
- Log the duration via `_logger?.LogInformation("Tool {ToolName} executed in {Duration}ms", toolName, stopwatch.ElapsedMilliseconds)`
- For tools that execute via GUI Proxy, the total duration includes queue wait time + execution time; both are captured by the overall Stopwatch

### How to verify
- [ ] Duration is measured using Stopwatch for precision (AC-03)
- [ ] Duration is reported in execution completed events (AC-02)
- [ ] Duration includes total time from start to completion (AC-03)

---

## TASK-006-11-04: Log GUI Proxy action queuing and completion events

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Subscribe to `IGUIProxy.RequestCompleted` event in `MCPViewModel` initialization
- On completion event: extract action name and duration from `GUIProxyResponse`
- If success: call `AddInfoLog($"GUI Proxy action completed: {response.RequestId} ({response.Duration?.TotalMilliseconds:F0}ms)", "GUIProxy")`
- If failure: call `AddWarningLog($"GUI Proxy action failed: {response.ErrorMessage}", "GUIProxy")`
- If timeout: call `AddErrorLog($"GUI Proxy action timed out: {response.RequestId}", "GUIProxy")`
- Subscribe to `IGUIProxy.ErrorOccurred` event
- On error: call `AddErrorLog($"GUI Proxy error: {ex.Message}", "GUIProxy")`
- These log entries provide visibility into the thread-safe COM operation pipeline
- Source field is set to "GUIProxy" to distinguish from direct tool execution logs

### How to verify
- [ ] GUI Proxy action completions are logged (AC-05)
- [ ] GUI Proxy failures and timeouts are logged (AC-05)
- [ ] Source field is "GUIProxy" for proxy events (AC-05)

---

## TASK-006-11-05: Ensure log entries are dispatched to the UI thread for ObservableCollection update

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify that the `AddLog()` method (from US-006-08) wraps the `ActivityLog.Add()` call in `Application.Current.Dispatcher.InvokeAsync()`
- If not already implemented, update `AddLog()` to check `Application.Current.Dispatcher.CheckAccess()`: if on UI thread, add directly; if not, use `Dispatcher.InvokeAsync` with `DispatcherPriority.Normal`
- This is critical because `ToolExecutionStarted`, `ToolExecutionCompleted`, and `ToolExecutionFailed` events fire from the MCP server background thread (MTA), not the GUI thread
- Similarly, `IGUIProxy.RequestCompleted` and `IGUIProxy.ErrorOccurred` events may fire from background threads
- Test scenario: multiple rapid tool executions should not cause `InvalidOperationException` on the ObservableCollection
- Consider batching log additions if rapid fire events cause performance issues (queue and flush every 100ms)

### How to verify
- [ ] Log entries from background threads are safely added to ObservableCollection (AC-06)
- [ ] No InvalidOperationException on rapid tool execution logging (AC-06)
- [ ] Activity log updates in real-time without manual refresh (AC-06)

---

## Dependency Graph
```
TASK-006-11-01 (MCPToolRegistry Execution Events)
       │
       └──▶ TASK-006-11-02 (ViewModel Event Handler)

TASK-006-11-03 (Stopwatch Duration Measurement)
       (independent)

TASK-006-11-04 (GUI Proxy Event Logging)
       (independent)

TASK-006-11-05 (UI Thread Dispatch Safety)
       (independent)

Tasks 01, 03, 04, and 05 are independent and can be developed in parallel.
Task 02 depends on 01 (events must exist to subscribe).
```
