# US-006-10: Restart Server

## User Story
**As a** System Admin,
**I want to** restart the MCP server,
**So that** I can recover from server issues without restarting the entire application.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: A "Restart" button is available in the server control panel
- [ ] AC-02: Clicking Restart stops the server, waits at least 1 second, then starts it again
- [ ] AC-03: The server status transitions through "Stopping..." -> "Stopped" -> "Starting..." -> "Running" during restart
- [ ] AC-04: All active SSE connections are closed during the stop phase before the server restarts
- [ ] AC-05: The Restart button is disabled while the server is already stopped (restart only applies when running)
- [ ] AC-06: If the stop phase succeeds but the start phase fails, the server remains in the stopped state with an error message
- [ ] AC-07: The restart operation is logged in the activity log with start and completion timestamps

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-003 | Page SHALL provide a Restart button that stops the server, waits briefly, and starts it again | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-10-01 | Implement RestartServerCommand in MCPViewModel that chains StopAsync, delay, and StartAsync | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-10-02 | Add Restart button to the Server Control panel XAML with CanExecute bound to IsServerRunning | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-10-03 | Implement MCPSSEServer.RestartAsync() with 1-second delay between stop and start (VR-006-007) | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | M |
| TASK-006-10-04 | Update ServerStatus property through intermediate states during restart | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-10-05 | Log restart operation start and completion in activity log | `ViewModels/MCPViewModel.cs` | S |

## Dependencies
- Depends on: US-006-01 (start MCP server), US-006-02 (stop MCP server) -- restart composes stop and start operations
- Blocks: None

## Notes
- The restart operation follows the Python pattern: `stop_server()` -> `time.sleep(1)` -> `start_server()`. In C# this is: `await StopAsync()` -> `await Task.Delay(1000)` -> `await StartAsync()`.
- VR-006-007 requires a minimum 1-second delay between stop and start to ensure the port is fully released by the OS.
- During the restart sequence, all buttons (Start, Stop, Restart, Test Connection) should be disabled to prevent conflicting operations.
- If the server was not running when Restart is clicked, the button should be disabled (CanExecute returns false).
- The restart is useful for recovering from issues like stale connections, tool registration errors, or protocol mismatches without closing the application.
