# US-007-01: Configure Odoo Connection

## User Story
**As a** CAD Engineer,
**I want to** configure the Odoo server URL, database, and credentials,
**So that** I can connect to the correct Odoo instance for my project.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays input fields for Odoo Server URL, Database Name, and Username
- [ ] AC-02: Settings Page displays a secure input field for Odoo API Token with masked display by default
- [ ] AC-03: API Token field has a show/hide toggle button that switches between PasswordBox and TextBox
- [ ] AC-04: Settings Page displays a connection timeout field with default value of 30 seconds
- [ ] AC-05: Server URL field validates format (must start with `http://` or `https://`) and shows inline error for invalid input
- [ ] AC-06: Database Name field validates non-empty, max 100 characters, alphanumeric/hyphens/underscores only
- [ ] AC-07: API Token field validates UUID format or non-empty string
- [ ] AC-08: Timeout field validates integer between 5 and 300 seconds inclusive
- [ ] AC-09: Clicking Save persists all Odoo connection settings to both `appsettings.json` and the `ServerConfigs` database table
- [ ] AC-10: API Token is persisted only to the `ServerConfigs` database table, never written to `appsettings.json` in plain text
- [ ] AC-11: Save button is disabled when any required field fails validation
- [ ] AC-12: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-001 | Settings Page SHALL provide input fields for Odoo Server URL, Database Name, and Username | Must |
| FR-007-002 | Settings Page SHALL provide a secure input field for Odoo API Token (masked display, show/hide toggle) | Must |
| FR-007-005 | Settings Page SHALL provide a connection timeout setting (in seconds, default 30) | Should |
| FR-007-006 | Settings Page SHALL persist all Odoo connection settings to both `appsettings.json` and the `ServerConfigs` database table | Must |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings and a Reset to Defaults button | Must |
| FR-007-027 | Unsaved changes SHALL be tracked with navigation confirmation dialog | Should |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-01-01 | Add Odoo connection input fields (Server URL, Database, Username, Token, Timeout) to Connection tab in XAML | `Views/Pages/SettingsPage.xaml` | M |
| TASK-007-01-02 | Implement PasswordBox/TextBox toggle for API Token with show/hide button | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-01-03 | Add ViewModel properties for Odoo connection fields (OdooServerUrl, OdooDatabaseName, OdooUsername, OdooApiToken, OdooTimeoutSeconds, IsTokenVisible) | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-01-04 | Implement INotifyDataErrorInfo validation for URL format, database name, token, and timeout range | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-01-05 | Implement SaveSettingsCommand that persists to appsettings.json (excluding token) and ServerConfigs DB table | `ViewModels/SettingsViewModel.cs` | L |
| TASK-007-01-06 | Extend ConfigurationLoader with SaveToJson method that serializes settings excluding sensitive fields | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M |
| TASK-007-01-07 | Implement HasUnsavedChanges tracking by comparing current values against snapshot of original values | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-01-08 | Add SyncLog audit entry on settings save with SyncType "settings_change" | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: US-008-01 (navigation framework must be in place for Settings Page to be accessible)
- Blocks: US-007-03 (Test Connection requires configured connection fields)

## Notes
- The Python application stores server config in `config/server.yaml` and token in `config/token.yaml`. The C# migration consolidates these into `appsettings.json` + EF Core `ServerConfigs` table.
- Token security is critical: the API token must never appear in `appsettings.json`. Use the `ServerConfigs` database table exclusively for token storage.
- The ViewModel should be registered as transient in DI so it loads fresh values each time the Settings Page is navigated to.
- ConfigKey mappings: `odoo.server_url`, `odoo.database`, `odoo.username`, `odoo.api_token`, `odoo.timeout_seconds`.
- Validation rules: VR-007-001 (URL format), VR-007-002 (URL non-empty), VR-007-003 (DB name format), VR-007-004 (token format), VR-007-005 (timeout range).
