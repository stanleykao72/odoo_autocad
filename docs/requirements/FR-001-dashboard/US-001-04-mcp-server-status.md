# US-001-04: MCP Server Status

## User Story
**As a** System Admin,
**I want to** see MCP server status,
**So that** I can verify AI assistant availability.

## Parent Feature
- **FR**: [FR-001-dashboard](../FR-001-dashboard/FR-001-dashboard.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Dashboard displays an MCP Server status card with running/stopped state
- [ ] AC-02: MCP Server status card displays the port number when the server is running (e.g., "Port: 8084")
- [ ] AC-03: Running state displays with a green indicator; stopped state displays with a red/gray indicator
- [ ] AC-04: Dashboard provides a toggle button to start/stop the MCP server
- [ ] AC-05: MCP server status updates in real-time when the server state changes
- [ ] AC-06: When MCP server fails to start, the card displays "MCP Server failed to start on port {port}. Port may be in use."
- [ ] AC-07: Toggle button label reflects current state (e.g., "Start" when stopped, "Stop" when running)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-001-003 | Dashboard SHALL display MCP Server status as a card with running/stopped state and port info | Should |
| FR-001-005 | Connected state SHALL display with green indicator; disconnected with red/gray indicator | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-001-04-01 | Create MCP Server status card in XAML alongside AutoCAD and Odoo cards, with icon, title, status indicator, port label, and toggle button | `Views/Pages/DashboardPage.xaml` | M |
| TASK-001-04-02 | Add ViewModel properties: IsMCPRunning (bool), MCPStatus (string), MCPPort (int) | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-04-03 | Implement ToggleMCPCommand using AsyncRelayCommand: calls StartServer() when stopped, StopServer() when running | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-04-04 | Implement async monitoring of MCP SSE server state with status change callback (matching Python's on_mcp_sse_status_update pattern) | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-04-05 | Add error handling for MCP server start failure: display error message in card and enable retry | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-04-06 | Bind toggle button content to IsMCPRunning to show "Start"/"Stop" dynamically, and apply green/red status indicator via BoolToColorConverter | `Views/Pages/DashboardPage.xaml` | S |
| TASK-001-04-07 | Integrate with MCP SSE Manager service for server lifecycle management (start, stop, get status) | `ViewModels/DashboardViewModel.cs` | L |

## Dependencies
- Depends on: US-001-01 (shares the same card layout pattern and BoolToColorConverter for status indicators)
- Depends on: MCP SSE Manager service interface being available for server lifecycle control
- Blocks: none

## Notes
- The Python implementation uses an SSE status indicator (green circle for running, red circle for stopped) with a callback `on_mcp_sse_status_update(message, is_running)`. The C# version should follow a similar event-driven pattern.
- The MCP Server status card is the third card in the horizontal row alongside AutoCAD and Odoo cards (see wireframe in FR document).
- FR-001-003 has priority "Should", so this user story is lower priority than the connection status cards (US-001-01) which are "Must".
- The MCP SSE Manager is part of a separate subsystem (MCP project). The ViewModel should depend on an abstracted interface rather than directly coupling to the MCP implementation.
- Consider using `DispatcherTimer` with 100ms interval for polling MCP server status, matching the Python GUI proxy pattern described in the FR document's implementation notes.
- Port information (default 8084) should be configurable and read from the application settings/configuration.
