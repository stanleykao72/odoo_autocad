# US-007-12: Import Configuration

## User Story
**As a** System Admin,
**I want to** import a configuration from a file,
**So that** I can restore settings on a new installation.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an "Import Configuration" button in the Advanced tab under Data Management section
- [ ] AC-02: Clicking "Import Configuration" opens a file open dialog with filter for JSON files (*.json)
- [ ] AC-03: The imported JSON file is validated as valid JSON format before applying
- [ ] AC-04: The imported JSON is validated for expected top-level keys: Application, Odoo, AutoCAD, MCP
- [ ] AC-05: If the JSON is not valid, an error dialog is shown: "The selected file is not a valid configuration file. Expected JSON format."
- [ ] AC-06: If the JSON is missing required sections, a warning is shown: "Configuration file is missing required sections: {missing_keys}. Import partially applied." and the available sections are still applied
- [ ] AC-07: After successful import, all settings fields in the ViewModel are updated with the imported values
- [ ] AC-08: Imported values are not automatically persisted; the user must click Save to persist them (import acts as "preview/load into form")
- [ ] AC-09: HasUnsavedChanges is set to true after a successful import
- [ ] AC-10: Import does not overwrite the API token unless explicitly present in the import file
- [ ] AC-11: Import operation is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-022 | Settings Page SHALL provide an "Import Configuration" button that loads settings from a user-chosen JSON file and applies them | Should |
| FR-007-027 | Unsaved changes SHALL be tracked with navigation confirmation dialog | Should |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-12-01 | Add "Import Configuration" button to Data Management section of Advanced tab in XAML | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-12-02 | Implement ImportConfigCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-12-03 | Implement IFileDialogService.ShowOpenFileDialog() for JSON file selection | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-12-04 | Implement JSON parsing and validation for structure and expected top-level keys | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M |
| TASK-007-12-05 | Map imported JSON values to ViewModel properties and set HasUnsavedChanges to true | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-12-06 | Display success, partial-import warning, or error feedback based on import result | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-12-07 | Log import operation to SyncLog table | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- Import acts as a "load into form" operation, not a direct database write. The user must explicitly click Save to persist imported values. This gives the user an opportunity to review and adjust values before committing.
- Validation rule: VR-007-013 (imported JSON must be valid JSON and contain expected top-level keys).
- The import should handle partial configuration files gracefully. If only the `Odoo` and `AutoCAD` sections are present, those should be applied and a warning shown for the missing `Application` and `MCP` sections.
- Security: If the imported file contains an `api_token` field, it should be applied to the ViewModel (so it appears in the form) but the user is still required to Save explicitly. This prevents accidental token overwrite.
- The `IFileDialogService` abstraction is used to enable testability. In production, it wraps the Win32 `OpenFileDialog`.
- Consider validating individual field values from the import against the same validation rules used for manual entry (VR-007-001 through VR-007-012).
