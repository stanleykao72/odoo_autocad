# TASKS: US-006-07 — Monitor Connections

> **Parent US**: [US-006-07](US-006-07-monitor-connections.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 4 | **Effort**: 4S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-006-01 (Start MCP Server - connections only exist when the server is running)
- [ ] US-006-03 (View Server Status - connection count is displayed in the status panel)
- [ ] `MCPSSEServer` exists with `ConcurrentDictionary<string, SSEConnection>` for connection tracking

## Acceptance Criteria
- [ ] AC-01: The page displays the current count of active SSE connections in the status panel
- [ ] AC-02: The connection count updates in real-time as clients connect and disconnect
- [ ] AC-03: When the server is stopped, the active connection count displays 0
- [ ] AC-04: The connection count accurately reflects the ConcurrentDictionary of tracked SSE connections in MCPSSEServer
- [ ] AC-05: SSE connection drops (client disconnects) are detected and the count is decremented accordingly

---

## TASK-006-07-01: Expose ActiveConnections count property from MCPSSEServer with change notification

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-006-07-02 |

### What to do
- Verify the existing `public int ActiveConnections => _connections.Count` property is in place
- Add an `event Action<int>? ConnectionCountChanged` event (if not already added in US-006-03)
- In `HandleSSEAsync`, after `_connections[connection.ConnectionId] = connection`, fire `ConnectionCountChanged?.Invoke(ActiveConnections)`
- In the `finally` block of `HandleSSEAsync`, after `_connections.TryRemove(connection.ConnectionId, out _)`, fire `ConnectionCountChanged?.Invoke(ActiveConnections)`
- In `StopAsync`, after `_connections.Clear()`, fire `ConnectionCountChanged?.Invoke(0)`
- Ensure the event fires on every connection state change (add, remove, clear) for accurate real-time tracking
- The `ConcurrentDictionary.Count` property is thread-safe, making direct reads safe from any thread

### How to verify
- [ ] ActiveConnections accurately reflects ConcurrentDictionary count (AC-04)
- [ ] ConnectionCountChanged fires on every connect/disconnect (AC-02)
- [ ] Count resets to 0 on server stop (AC-03)

---

## TASK-006-07-02: Bind ActiveConnectionCount ViewModel property to MCPSSEServer.ActiveConnections

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-07-01 |
| Blocks | None |

### What to do
- Subscribe to `MCPSSEServer.ConnectionCountChanged` event in `MCPViewModel` initialization
- Create handler `OnConnectionCountChanged(int count)`: use `Dispatcher.InvokeAsync()` to update `ActiveConnectionCount = count` on the UI thread
- When server starts, initialize `ActiveConnectionCount = _mcpServer.ActiveConnections`
- When server stops, set `ActiveConnectionCount = 0`
- Add an activity log entry on connection changes: "SSE connection established: {connectionId}" on connect, "SSE connection closed: {connectionId}" on disconnect (requires enhancing the event to include connection ID, or subscribing to separate connect/disconnect events)
- Alternatively, poll `_mcpServer.ActiveConnections` from the existing status update mechanism if event-based approach is not feasible

### How to verify
- [ ] ActiveConnectionCount updates in real-time on connect/disconnect (AC-02)
- [ ] Count reflects accurate ConcurrentDictionary state (AC-04)
- [ ] Connection events are logged in activity log (AC-02)

---

## TASK-006-07-03: Display connection count in the Server Control panel XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Server Control panel (created in US-006-03), add or verify the "Connections:" label with value TextBlock bound to `{Binding ActiveConnectionCount}`
- Ensure the binding is present in the status info grid alongside Port, Uptime, Protocol, and Tools labels
- Apply numeric formatting if needed (though integer values are simple)
- Add a `ToolTip` on the connection count: "Number of AI assistant clients currently connected via SSE"
- When the count is 0 and server is stopped, the value should display "0" (not blank or hidden)

### How to verify
- [ ] Connection count is displayed with label "Connections: {count}" (AC-01)
- [ ] Display shows 0 when server is stopped (AC-03)

---

## TASK-006-07-04: Update connection count on SSE connect/disconnect events using the status callback

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add separate events for more detailed connection tracking: `event Action<string>? ClientConnected` (fires with connectionId), `event Action<string>? ClientDisconnected` (fires with connectionId)
- In `HandleSSEAsync`, after creating the connection and adding to `_connections`, fire `ClientConnected?.Invoke(connection.ConnectionId)`
- In the `finally` block, after removing from `_connections`, fire `ClientDisconnected?.Invoke(connection.ConnectionId)`
- These detailed events allow the ViewModel to log individual connection events with IDs
- The heartbeat mechanism (30-second interval, FR-006-019) already detects stale connections via `OperationCanceledException` in the SSE loop, which triggers the `finally` block and thus `ClientDisconnected`
- Ensure connection cleanup in `StopAsync` also fires disconnect events for each connection being closed

### How to verify
- [ ] Client connect events fire with connectionId (AC-02)
- [ ] Client disconnect events fire on clean and stale disconnections (AC-05)
- [ ] Heartbeat detects dropped connections and triggers cleanup (AC-05)

---

## Dependency Graph
```
TASK-006-07-01 (MCPSSEServer ConnectionCountChanged Event)
       │
       └──▶ TASK-006-07-02 (ViewModel Binding to Event)

TASK-006-07-03 (XAML Display)
       (independent)

TASK-006-07-04 (Detailed Connect/Disconnect Events)
       (independent)

Tasks 01, 03, and 04 are independent.
Task 02 depends on 01 (event must exist to subscribe).
```
