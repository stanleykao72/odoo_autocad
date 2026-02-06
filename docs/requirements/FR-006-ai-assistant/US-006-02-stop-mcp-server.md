# US-006-02: Stop MCP Server

## User Story
**As a** CAD Engineer,
**I want to** stop the MCP SSE server,
**So that** I can free the port and resources when AI assistance is not needed.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: Clicking the "Stop Server" button shuts down the MCP SSE server and releases the port
- [ ] AC-02: All active SSE connections are gracefully closed before the server shuts down
- [ ] AC-03: The Stop button toggles back to "Start Server" once the server has stopped
- [ ] AC-04: The Stop button is disabled while the server is already stopped
- [ ] AC-05: The status indicator changes from green (Running) to red (Stopped) upon successful stop
- [ ] AC-06: The DispatcherTimer for GUI Proxy polling is stopped when the server stops
- [ ] AC-07: If the server is already stopped, the stop operation is idempotent and returns without error

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-002 | Page SHALL provide a Stop button to shut down the MCP SSE server and close all SSE connections | Must |
| FR-006-007 | Server stop SHALL gracefully close all active SSE connections before shutting down | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-02-01 | Implement StopServerCommand in MCPViewModel with graceful shutdown logic | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-02-02 | Implement MCPSSEServer.StopAsync() to close all SSE connections and release port | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | M |
| TASK-006-02-03 | Stop DispatcherTimer and clean up IGUIProxy resources on server stop | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | S |
| TASK-006-02-04 | Update XAML button binding to show "Start Server" when stopped and disable Stop while stopped | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-02-05 | Log server shutdown event and connection closures in activity log | `ViewModels/MCPViewModel.cs` | S |

## Dependencies
- Depends on: US-006-01 (start MCP server - server must be running to stop it)
- Blocks: None

## Notes
- Graceful shutdown means the server sends a close event to all connected SSE clients before terminating the HTTP listener.
- The ConcurrentDictionary of active SSE connections must be cleared during stop.
- If the server is already stopped, StopAsync should be idempotent and return true (VR-006-004).
- The stop operation should log the number of connections that were closed during shutdown.
