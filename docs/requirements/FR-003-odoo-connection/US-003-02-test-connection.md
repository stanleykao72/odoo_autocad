# US-003-02: Test Connection

## User Story
**As a** CAD Engineer,
**I want to** test the Odoo connection before saving settings,
**So that** I know the credentials are valid before proceeding.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A "Test Connection" button is visible on the connection settings panel
- [ ] AC-02: Clicking "Test Connection" calls `TestConnectionAsync()` without persisting the session
- [ ] AC-03: On success, a green status message displays "Connection test successful" with server version info
- [ ] AC-04: On authentication failure (uid = false), the message displays "Authentication failed. Please verify your database name, username, and password."
- [ ] AC-05: On network failure (HttpRequestException), the message displays "Unable to reach Odoo server at {ServerUrl}. Please check your network connection and server URL."
- [ ] AC-06: On timeout (TaskCanceledException), the message displays "Connection timed out after {TimeoutSeconds} seconds."
- [ ] AC-07: On SSL/TLS error, the message displays "SSL certificate validation failed for {ServerUrl}."
- [ ] AC-08: The button is disabled while the test is in progress (shows a progress indicator)
- [ ] AC-09: The button is disabled when required credential fields are empty or invalid
- [ ] AC-10: The test operation does not modify the IsConnected state or establish a persistent session

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-002 | The page SHALL provide a "Test Connection" button that validates credentials without persisting the session | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-02-01 | Add "Test Connection" button to the connection settings button bar in XAML | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-02-02 | Implement TestConnectionCommand as IAsyncRelayCommand calling IOdooService.TestConnectionAsync() | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-02-03 | Implement TestConnectionAsync() in OdooService to validate credentials via `/web/session/authenticate` without retaining session | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-02-04 | Define TestConnectionAsync(serverUrl, database, username, password) in IOdooService interface | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |
| TASK-003-02-05 | Add error handling with user-friendly messages for each failure scenario (network, auth, timeout, SSL) | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-02-06 | Bind button IsEnabled to validation state and IsRunning property for progress indication | `Views/Pages/OdooConnectionPage.xaml` | S |

## Dependencies
- Depends on: US-003-01
- Blocks: None

## Notes
- The "Test Connection" operation should be lightweight: authenticate, read server version, then discard the session without storing it in the service state.
- Error messages should map to the Error Handling table in the FR document (Section 9).
- The Python implementation does not have an explicit test connection; it simply calls `connect_odoo()`. The C# version adds this as a separate validation step for better UX.
- Timeout is governed by `Odoo:TimeoutSeconds` from appsettings.json (default: 30 seconds).
