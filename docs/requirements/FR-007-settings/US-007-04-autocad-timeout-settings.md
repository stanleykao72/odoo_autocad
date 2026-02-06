# US-007-04: AutoCAD Timeout Settings

## User Story
**As a** CAD Engineer,
**I want to** adjust AutoCAD connection timeout and retry attempts,
**So that** I can handle slower AutoCAD startup on my machine.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a read-only field for AutoCAD COM ProgId showing "AutoCAD.Application" in the AutoCAD tab
- [ ] AC-02: Settings Page displays a configurable connection timeout field with default value of 10 seconds
- [ ] AC-03: Settings Page displays a configurable retry attempts field with default value of 3
- [ ] AC-04: Connection timeout validates as integer between 5 and 120 seconds inclusive, with inline error for out-of-range values
- [ ] AC-05: Retry attempts validates as integer between 1 and 10 inclusive, with inline error for out-of-range values
- [ ] AC-06: Saving persists AutoCAD settings to `ServerConfigs` database table with keys `autocad.prog_id`, `autocad.connection_timeout`, `autocad.retry_attempts`
- [ ] AC-07: AutoCAD settings are loaded from database on page initialization with fallback to defaults
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-007 | Settings Page SHALL display the AutoCAD COM ProgId (default: `AutoCAD.Application`, read-only unless advanced mode) | Should |
| FR-007-008 | Settings Page SHALL provide a configurable connection timeout in seconds (default: 10) | Should |
| FR-007-009 | Settings Page SHALL provide a configurable retry attempt count (default: 3, range 1-10) | Should |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-04-01 | Add AutoCAD tab with ProgId (read-only), Connection Timeout, and Retry Attempts fields in XAML | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-04-02 | Add AutoCADProgId, AutoCADConnectionTimeout, and AutoCADRetryAttempts properties to ViewModel | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-04-03 | Implement validation for timeout (5-120) and retry attempts (1-10) with inline error messages | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-04-04 | Add numeric input constraints to XAML fields (integer-only input, spinner controls) | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-04-05 | Persist AutoCAD settings to ServerConfigs table on Save | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-04-06 | Load AutoCAD settings from ServerConfigs table on page initialization with default fallback | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The AutoCAD COM ProgId is `AutoCAD.Application` and is read-only in normal mode. An advanced mode toggle could make it editable in future iterations, but for v6.0 it remains fixed.
- In the Python application, AutoCAD connection settings were hardcoded. The C# migration makes them configurable through the Settings Page.
- Validation rules: VR-007-007 (retry attempts 1-10), VR-007-008 (timeout 5-120 seconds).
- ConfigKey mappings: `autocad.prog_id`, `autocad.connection_timeout`, `autocad.retry_attempts`.
