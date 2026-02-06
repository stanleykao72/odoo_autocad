# US-006-04: Test Connection

## User Story
**As a** System Admin,
**I want to** test the MCP connection,
**So that** I can verify the server is responding correctly to JSON-RPC requests.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: A "Test Connection" button is available in the server control panel
- [ ] AC-02: Clicking the button sends a test_connection tool call to the running MCP server
- [ ] AC-03: The test result displays success or failure status with a timestamp
- [ ] AC-04: The test verifies both the HTTP endpoint (/health) and the JSON-RPC protocol layer
- [ ] AC-05: The Test Connection button is disabled when the server is not running
- [ ] AC-06: A timeout message "Connection test timed out. Server may be unresponsive." is shown if the test does not complete within the timeout period
- [ ] AC-07: A failure message "Connection test failed: {error}. Server may not be responding to JSON-RPC requests." is shown on error

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-031 | Page SHALL provide a "Test Connection" button that sends a test_connection tool call to the server | Must |
| FR-006-032 | Test result SHALL display success/failure status with timestamp | Must |
| FR-006-033 | Test SHALL verify both the HTTP endpoint (/health) and the JSON-RPC protocol layer | Should |
| FR-006-034 | Test button SHALL be disabled when the server is not running | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-04-01 | Implement TestConnectionCommand in MCPViewModel that calls health and JSON-RPC endpoints | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-04-02 | Add Test Connection button to XAML with CanExecute bound to IsServerRunning | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-04-03 | Implement two-phase test: HTTP GET /health check followed by JSON-RPC test_connection tool call via POST /messages | `ViewModels/MCPViewModel.cs` | M |
| TASK-006-04-04 | Display test results (TestConnectionResult, IsTestConnectionSuccess, LastTestTime) in the status panel | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-04-05 | Handle timeout and failure scenarios with appropriate user-facing error messages | `ViewModels/MCPViewModel.cs` | S |

## Dependencies
- Depends on: US-006-01 (start MCP server - server must be running to test connection)
- Blocks: None

## Notes
- The two-phase test approach first checks that the HTTP transport layer is responding (GET /health returns 200 with valid JSON), then verifies the MCP protocol layer is functioning (POST /messages with a JSON-RPC 2.0 tools/call request for test_connection returns a valid response).
- The test uses HttpClient to connect to http://localhost:{port}/health and http://localhost:{port}/messages.
- Validation rule VR-006-006 enforces that the Test Connection button is disabled when the server is not running.
- The test result area shows: success/failure icon, timestamp of last test, and response details or error message.
