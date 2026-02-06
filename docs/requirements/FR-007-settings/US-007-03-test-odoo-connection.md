# US-007-03: Test Odoo Connection

## User Story
**As a** CAD Engineer,
**I want to** test the Odoo connection from the settings page,
**So that** I can verify my settings are correct before starting work.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a "Test Connection" button in the Connection tab
- [ ] AC-02: Clicking "Test Connection" calls `IOdooService.TestConnectionAsync()` with the current URL, database, username, and token values
- [ ] AC-03: While testing, a loading indicator is displayed and the button is disabled (IsTestingConnection = true)
- [ ] AC-04: On success, a green success message "Connected successfully" is displayed inline next to the button
- [ ] AC-05: On timeout failure, the message "Connection test timed out after {timeout} seconds. Check the server URL and network." is displayed
- [ ] AC-06: On authentication failure, the message "Authentication failed. Please verify your API token and username." is displayed and the token field is highlighted
- [ ] AC-07: On network error, the message "Cannot reach server at {url}. Check network connectivity and firewall settings." is displayed
- [ ] AC-08: The test result (ConnectionTestResult and ConnectionTestMessage) is displayed until the next test or until connection fields are modified
- [ ] AC-09: Test Connection button is disabled when required connection fields are empty or invalid

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-003 | Settings Page SHALL provide a "Test Connection" button that validates the Odoo server URL and credentials by calling `IOdooService.TestConnectionAsync()` | Must |
| FR-007-004 | Test Connection SHALL display success or failure result inline with specific error details (timeout, auth failure, network error) | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-03-01 | Add "Test Connection" button with inline result display area to Connection tab XAML | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-03-02 | Implement TestConnectionCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-03-03 | Add IsTestingConnection, ConnectionTestResult (bool?), and ConnectionTestMessage properties to ViewModel | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-03-04 | Implement IOdooService.TestConnectionAsync(url, database, username, token) method | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | M |
| TASK-007-03-05 | Add error classification logic to differentiate timeout, auth failure, and network errors | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-03-06 | Add XAML data triggers for success (green), failure (red), and testing (spinner) states | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-03-07 | Clear test result when connection fields are modified (reset ConnectionTestResult to null) | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: US-007-01 (connection fields must be configured before testing)
- Blocks: None

## Notes
- The test connection operation should use the timeout value currently entered in the settings form, not the previously saved value.
- Error classification should distinguish between: `TaskCanceledException` / `TimeoutException` (timeout), `HttpRequestException` with 401/403 status (auth failure), and other `HttpRequestException` / `SocketException` (network error).
- The `ConnectionTestResult` property uses nullable bool: `null` = not tested, `true` = success, `false` = failure. This allows XAML to show three distinct visual states.
- The test should be cancellable if the user navigates away from the settings page.
