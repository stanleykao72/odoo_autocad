# US-007-11: Export Configuration

## User Story
**As a** System Admin,
**I want to** export the current configuration to a file,
**So that** I can back up or share settings across machines.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an "Export Configuration" button in the Advanced tab under Data Management section
- [ ] AC-02: Clicking "Export Configuration" opens a file save dialog with default filter for JSON files (*.json)
- [ ] AC-03: The exported JSON file contains all current settings organized under top-level keys: Application, Odoo, AutoCAD, MCP
- [ ] AC-04: The exported JSON excludes sensitive fields (API token is not included in the export)
- [ ] AC-05: On successful export, a success message is displayed with the export file path
- [ ] AC-06: If export fails due to file permissions or disk error, an error dialog is shown: "Failed to export configuration to {path}. Check file permissions."
- [ ] AC-07: The export operation does not modify any current settings
- [ ] AC-08: Export operation is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-021 | Settings Page SHALL provide an "Export Configuration" button that saves current settings to a user-chosen JSON file | Should |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-11-01 | Add "Export Configuration" button to Data Management section of Advanced tab in XAML | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-11-02 | Implement ExportConfigCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-11-03 | Implement IFileDialogService.ShowSaveFileDialog() for JSON file selection | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-11-04 | Serialize current settings to JSON excluding sensitive fields (token) | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M |
| TASK-007-11-05 | Write serialized JSON to the user-selected file path with error handling | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-11-06 | Display success or error feedback after export operation | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-11-07 | Log export operation to SyncLog table | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The exported JSON structure should match the `appsettings.json` format with top-level keys: `Application`, `Odoo`, `AutoCAD`, `MCP`. This ensures compatibility with the Import Configuration feature (US-007-12).
- Security: The API token (`odoo.api_token`) must be explicitly excluded from the export. The exported Odoo section should include server URL, database, username, and timeout, but not the token.
- The `IFileDialogService` abstraction is used to enable testability. In production, it wraps the Win32 `SaveFileDialog`.
- The export includes both `ServerConfigs` and `UserPreferences` values, merged into the appropriate JSON sections.
- Consider including a metadata section in the export with export timestamp, application version, and source machine name for traceability.
