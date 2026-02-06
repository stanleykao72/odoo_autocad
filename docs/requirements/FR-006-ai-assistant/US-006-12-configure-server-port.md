# US-006-12: Configure Server Port

## User Story
**As a** System Admin,
**I want to** configure the server port,
**So that** I can avoid port conflicts with other applications.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: The page displays the currently configured server port (default: 8084)
- [ ] AC-02: The port field is editable when the server is stopped
- [ ] AC-03: The port field is disabled (read-only) when the server is running
- [ ] AC-04: Port input is validated to be an integer between 1024 and 65535
- [ ] AC-05: Invalid port values display a validation error message and prevent the server from starting
- [ ] AC-06: Changing the port automatically updates all displayed endpoint URLs (SSE, Health, Messages)
- [ ] AC-07: The configured port is persisted so it survives application restarts

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-045 | Page SHALL display the configured server port (default: 8084) | Must |
| FR-006-046 | Page SHALL allow the user to change the server port when the server is stopped | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-12-01 | Add editable port TextBox to Configuration panel XAML with IsEnabled bound to !IsServerRunning | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-12-02 | Implement port validation in MCPViewModel (VR-006-001: range 1024-65535, integer only) | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-12-03 | Update computed endpoint URL properties (SSEEndpointUrl, HealthEndpointUrl) when ServerPort changes | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-12-04 | Persist port configuration to application settings or configuration file | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-12-05 | Add port availability check on server start (VR-006-002) and display error if port is in use | `OdooAutoCAD.MCP/Server/MCPSSEServer.cs` | M |
| TASK-006-12-06 | Display validation error styling (red border, error message) for invalid port input | `Views/Pages/MCPAssistantPage.xaml` | S |

## Dependencies
- Depends on: None (port configuration is available at all times)
- Blocks: US-006-01 (start MCP server - server uses the configured port)

## Notes
- Validation rules:
  - VR-006-001: Port must be an integer between 1024 and 65535.
  - VR-006-002: Port must not be in use by another process before starting.
  - VR-006-005: Port configuration field must be disabled while the server is running.
- The default port is 8084, matching the Python implementation's MCPSSEManager default.
- When the port changes, the following computed properties must be recalculated:
  - SSEEndpointUrl: `http://localhost:{port}/sse`
  - HealthEndpointUrl: `http://localhost:{port}/health`
  - Root URL: `http://localhost:{port}/`
  - Messages URL: `http://localhost:{port}/messages`
- Port persistence can use the application's existing configuration management system (YAML or settings store).
- The port field should use numeric input filtering to prevent non-numeric characters from being entered.
