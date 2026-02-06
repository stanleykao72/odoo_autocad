# US-003-01: Enter Odoo Credentials

## User Story
**As a** CAD Engineer,
**I want to** enter Odoo server URL and credentials on a settings form,
**So that** I can connect to our company's Odoo instance.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The page displays input fields for Server URL, Database Name, Username, and Password/API Key
- [ ] AC-02: The Password field uses masked input (PasswordBox) so characters are not visible
- [ ] AC-03: Server URL, Database, and Username fields are pre-populated from `appsettings.json` on launch
- [ ] AC-04: Server URL validates as a valid HTTPS URL pattern (`^https?://[^\s/]+`); HTTP triggers a warning but is not blocked
- [ ] AC-05: Server URL has trailing slashes stripped before use (matching `OdooService.ConnectAsync` behavior)
- [ ] AC-06: Database Name field is not empty and contains only alphanumeric characters, hyphens, and underscores
- [ ] AC-07: Username field is not empty when using session-based authentication
- [ ] AC-08: Password field is not empty with a minimum length of 1 character
- [ ] AC-09: "Connect" and "Test Connection" buttons are disabled when Server URL is empty
- [ ] AC-10: The Password/API Key is never stored in `appsettings.json` in plaintext

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-001 | The page SHALL provide input fields for Server URL, Database Name, Username, and Password/API Key | Must |
| FR-003-008 | Connection settings SHALL be pre-populated from `appsettings.json` on application launch | Must |
| FR-003-009 | The Password/API Key field SHALL use PasswordBox (masked input) and SHALL NOT be stored in plaintext | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-01-01 | Create Connection Settings GroupBox with Grid layout containing 4 labeled input rows (Server URL, Database, Username, Password) | `Views/Pages/OdooConnectionPage.xaml` | M |
| TASK-003-01-02 | Add ViewModel properties for ServerUrl, Database, Username, Password with change notification | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-01-03 | Implement INotifyDataErrorInfo validation for all credential fields (URL pattern, non-empty, allowed characters) | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-01-04 | Load default values from IConfiguration Odoo section on ViewModel initialization | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-01-05 | Add Odoo configuration section to appsettings.json with ServerUrl, Database, Username, TimeoutSeconds | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | S |
| TASK-003-01-06 | Bind Connect/Test Connection button IsEnabled to validation state (all required fields non-empty and valid) | `Views/Pages/OdooConnectionPage.xaml` | S |

## Dependencies
- Depends on: None
- Blocks: US-003-02, US-003-03, US-003-04, US-003-12

## Notes
- The C# implementation uses session-based JSON-RPC authentication (`/web/session/authenticate`) instead of the Python token-based Basic Auth approach.
- Configuration is loaded from `appsettings.json` (`Odoo:ServerUrl`, `Odoo:Database`, `Odoo:Username`) instead of YAML files.
- The Password is intentionally excluded from `appsettings.json` for security; see US-003-12 for secure credential persistence.
- Validation rules VR-003-001 through VR-003-006 from the FR document apply to this story.
