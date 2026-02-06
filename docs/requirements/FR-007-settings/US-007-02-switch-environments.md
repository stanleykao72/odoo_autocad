# US-007-02: Switch Environments

## User Story
**As a** CAD Engineer,
**I want to** switch between different environment configurations (staging, production),
**So that** I can work against the right server without manual editing.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays an environment selector dropdown with options: Development, Staging, Production
- [ ] AC-02: The currently active environment is pre-selected in the dropdown on page load
- [ ] AC-03: Selecting a new environment loads the corresponding settings from `appsettings.{Environment}.json`
- [ ] AC-04: A confirmation dialog appears before overwriting current values when switching environments
- [ ] AC-05: Confirmation dialog shows which settings will change (server URL, database name, etc.)
- [ ] AC-06: Cancelling the confirmation dialog reverts the dropdown to the previous environment selection
- [ ] AC-07: After confirming environment switch, all connection fields update to reflect the new environment values
- [ ] AC-08: Environment selection is persisted to `UserPreferences` table with key `app.environment`
- [ ] AC-09: Environment switch is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-023 | Settings Page SHALL provide an environment selector for switching between pre-defined configurations (Development, Staging, Production) | Should |
| FR-007-024 | Switching environments SHALL load the corresponding settings and prompt the user to confirm before overwriting current values | Should |
| FR-007-027 | Unsaved changes SHALL be tracked with navigation confirmation dialog | Should |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-02-01 | Add environment selector dropdown to Connection tab in XAML bound to SelectedEnvironment and AvailableEnvironments | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-02-02 | Implement SelectedEnvironment property, AvailableEnvironments collection, and ChangeEnvironmentCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-02-03 | Implement GetAvailableConfigs() in ConfigurationLoader to discover available appsettings.{env}.json files | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M |
| TASK-007-02-04 | Implement confirmation dialog logic showing which settings will change and handling cancel/confirm | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-02-05 | Load environment-specific values from appsettings.{Environment}.json and populate ViewModel fields | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-02-06 | Persist selected environment to UserPreferences table with key `app.environment` | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: US-007-01 (connection fields must exist for environment values to populate them)
- Blocks: None

## Notes
- The Python application uses multiple YAML files (`server_khs.yaml`, `server_prod.yaml`, `server_prod_khs.yaml`) selected via `token.yaml > server_file`. The C# migration uses the standard .NET `appsettings.{Environment}.json` pattern.
- The environment selector populates from discovered `appsettings.*.json` files in the application directory.
- The load order on startup is: `appsettings.json` (base) -> `appsettings.{Environment}.json` (override) -> `ServerConfigs` DB (override) -> `UserPreferences` DB (override).
- ConfigKey mapping: `app.environment` stored in `UserPreferences` table.
