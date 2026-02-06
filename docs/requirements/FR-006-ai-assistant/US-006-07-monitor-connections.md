# US-006-07: Monitor Connections

## User Story
**As a** CAD Engineer,
**I want to** monitor active SSE connections,
**So that** I can see how many AI assistant clients are currently connected.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: The page displays the current count of active SSE connections in the status panel
- [ ] AC-02: The connection count updates in real-time as clients connect and disconnect
- [ ] AC-03: When the server is stopped, the active connection count displays 0
- [ ] AC-04: The connection count accurately reflects the ConcurrentDictionary of tracked SSE connections in MCPSSEServer
- [ ] AC-05: SSE connection drops (client disconnects) are detected and the count is decremented accordingly

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-026 | Page SHALL display the number of active SSE connections | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-07-01 | Expose ActiveConnections count property from MCPSSEServer with change notification | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | S |
| TASK-006-07-02 | Bind ActiveConnectionCount ViewModel property to MCPSSEServer.ActiveConnections | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-07-03 | Display connection count in the Server Control panel XAML with label "Connections: {count}" | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-07-04 | Update connection count on SSE connect/disconnect events using the status callback | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | S |

## Dependencies
- Depends on: US-006-01 (start MCP server - connections only exist when the server is running)
- Blocks: None

## Notes
- Active connections are tracked in a ConcurrentDictionary<string, SSEConnection> inside MCPSSEServer. The count is simply the dictionary's Count property.
- Each new SSE connection triggers a "connected" event with a connection ID (FR-006-020). Disconnections are detected when the SSE response stream is terminated or the heartbeat detects a closed connection.
- The heartbeat mechanism (every 30 seconds, FR-006-019) helps detect stale connections that may not have cleanly disconnected.
- The connection count is displayed in the Server Control panel alongside the status indicator, port, and tool count.
