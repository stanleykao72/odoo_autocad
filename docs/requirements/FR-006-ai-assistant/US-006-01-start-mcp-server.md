# US-006-01: Start MCP Server

## User Story
**As a** CAD Engineer,
**I want to** start the MCP SSE server from the UI,
**So that** AI assistants can connect and help me with AutoCAD operations.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: Clicking the "Start Server" button launches the MCP SSE server on the configured port (default 8084)
- [ ] AC-02: The Start button toggles to "Stop Server" once the server is running
- [ ] AC-03: The Start button is disabled while the server is already running
- [ ] AC-04: Server start verifies port availability before attempting to bind; if the port is in use, an error message is displayed: "Port {port} is already in use. Please choose a different port or stop the conflicting application."
- [ ] AC-05: All MCP tools are registered with the tool registry before the server accepts connections
- [ ] AC-06: The server gracefully handles start failure by remaining in the stopped state and logging the error
- [ ] AC-07: The status indicator changes from red (Stopped) to green (Running) upon successful start
- [ ] AC-08: The server exposes GET /sse, POST /messages, GET /health, and GET / endpoints after startup

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-001 | Page SHALL provide a Start button to launch the MCP SSE server on the configured port | Must |
| FR-006-004 | Start/Stop button SHALL toggle appearance based on server state | Must |
| FR-006-005 | Server start SHALL verify port availability before attempting to bind | Must |
| FR-006-006 | Server start SHALL register all MCP tools with the tool registry before accepting connections | Must |
| FR-006-007 | Server stop SHALL gracefully close all active SSE connections before shutting down | Must |
| FR-006-008 | Page SHALL disable the Start button while the server is already running and disable Stop while already stopped | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-01-01 | Implement StartServerCommand in MCPViewModel with port validation and error handling | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-01-02 | Add Start button with toggle binding and disabled state in XAML | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-01-03 | Implement MCPSSEServer.StartAsync() with port availability check and endpoint registration | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | L |
| TASK-006-01-04 | Register all MCP tools via MCPToolRegistry.RegisterAllTools() during server startup | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` | M |
| TASK-006-01-05 | Initialize IGUIProxy and start DispatcherTimer (100ms) for proxy polling on server start | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | M |
| TASK-006-01-06 | Wire status callback to update ViewModel properties (IsServerRunning, ServerStatus) on state change | `ViewModels/MCPViewModel.cs` | S |

## Dependencies
- Depends on: US-008-06 (app lifecycle - application must be initialized before MCP server can start)
- Blocks: US-006-02 (stop server), US-006-04 (test connection), US-006-07 (monitor connections), US-006-10 (restart server), US-006-11 (view tool execution)

## Notes
- The server runs on a background thread (MTA) using Task.Run with WebApplication.RunAsync. The DispatcherTimer on the GUI thread polls the IGUIProxy queue every 100ms to execute AutoCAD COM operations on the STA thread.
- Port validation must check both the configured range (1024-65535, per VR-006-001) and runtime availability (VR-006-002).
- If the server is already running, StartAsync should be idempotent and return true without error (VR-006-003).
- The server must report MCP protocol version 2024-11-05 during the initialize handshake.
