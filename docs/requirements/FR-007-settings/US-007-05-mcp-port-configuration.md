# US-007-05: MCP Port Configuration

## User Story
**As a** System Admin,
**I want to** configure the MCP SSE server port number,
**So that** I can avoid port conflicts with other services on the machine.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an input field for MCP SSE Server port number in the MCP tab with default value 8084
- [ ] AC-02: Port number validates as integer between 1024 and 65535 inclusive, with inline error for out-of-range values
- [ ] AC-03: A warning is displayed (non-blocking) if the entered port appears to be in use, with message "Port {port} appears to be in use. MCP server may not start. Choose a different port."
- [ ] AC-04: Saving persists the port number to `ServerConfigs` database table with key `mcp.port`
- [ ] AC-05: The port value is also saved to the `appsettings.json` MCP section
- [ ] AC-06: Port setting is loaded from database on page initialization with fallback to default 8084
- [ ] AC-07: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-010 | Settings Page SHALL provide an input field for MCP SSE Server port number (default: 8084) | Must |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-05-01 | Add MCP tab with SSE Server Port input field in XAML | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-05-02 | Add MCPPort property to ViewModel with default value 8084 | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-05-03 | Implement validation for port range (1024-65535) with inline error messages | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-05-04 | Implement optional port-in-use detection and display non-blocking warning | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-05-05 | Persist MCP port to ServerConfigs table and appsettings.json on Save | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-05-06 | Load MCP port from ServerConfigs table on page initialization with fallback to default | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The Python application uses `--port 8084` as a CLI argument or defaults to port 8084 in `util_mcp_sse_manager.py`. The C# migration makes this configurable through the Settings Page.
- Port-in-use detection can be implemented by attempting a TCP connection to localhost on the specified port. This check should be non-blocking and only provides a warning, not a hard validation failure.
- Validation rule: VR-007-006 (port range 1024-65535, must not conflict with well-known ports).
- ConfigKey mapping: `mcp.port` in `ServerConfigs` table.
