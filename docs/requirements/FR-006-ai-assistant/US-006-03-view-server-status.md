# US-006-03: View Server Status

## User Story
**As a** CAD Engineer,
**I want to** see the current server status (Running/Stopped),
**So that** I know whether AI assistants can connect.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: The page displays a visual status indicator (green ellipse for Running, red ellipse for Stopped)
- [ ] AC-02: The current server port number is displayed in the status panel
- [ ] AC-03: The number of active SSE connections is displayed and updated in real-time
- [ ] AC-04: The number of registered MCP tools is displayed in the status panel
- [ ] AC-05: The MCP session initialized state (whether a client has completed the handshake) is displayed
- [ ] AC-06: Status updates are pushed to the UI in real-time via a callback mechanism when server state changes
- [ ] AC-07: Server uptime is displayed when the server is running (formatted as Xh Ym Zs)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-024 | Page SHALL display server running state with a visual indicator (green for running, red for stopped) | Must |
| FR-006-025 | Page SHALL display the current server port number | Must |
| FR-006-026 | Page SHALL display the number of active SSE connections | Should |
| FR-006-027 | Page SHALL display the number of registered tools | Should |
| FR-006-028 | Page SHALL display whether the MCP session is initialized | Should |
| FR-006-029 | Status SHALL update in real-time via a callback mechanism when server state changes | Must |
| FR-006-030 | Page SHALL display server uptime when running | Could |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-03-01 | Create Server Control panel XAML layout with status indicator Ellipse, port label, uptime, connections, and tool count | `Views/Pages/MCPAssistantPage.xaml` | M |
| TASK-006-03-02 | Implement BoolToColorConverter for mapping IsServerRunning to green/red Ellipse fill | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-03-03 | Add ViewModel properties: IsServerRunning, ServerStatus, ServerPort, ActiveConnectionCount, RegisteredToolCount, IsSessionInitialized, ServerUptime | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-03-04 | Implement status callback handler that updates ViewModel properties on server state changes | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-03-05 | Implement uptime timer that updates ServerUptime property every second while server is running | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-03-06 | Expose ActiveConnections count and IsInitialized from MCPSSEServer for ViewModel binding | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | S |

## Dependencies
- Depends on: None (status panel is always visible regardless of server state)
- Blocks: None

## Notes
- The status panel is the top section of the MCP Assistant page and is always visible.
- The Ellipse fill uses a BoolToColorConverter: true maps to green (#4CAF50), false maps to red (#F44336).
- Uptime is calculated from ServerStartTime and formatted as "Xh Ym Zs". When the server is stopped, uptime displays "--" or is hidden.
- The ActiveConnectionCount reflects the size of the ConcurrentDictionary tracking SSE connections in MCPSSEServer.
- The status callback pattern matches the Python `set_status_callback(callback)` approach but uses C# events or INotifyPropertyChanged.
