# US-007-09: Change Log Level

## User Story
**As a** System Admin,
**I want to** change the logging level for troubleshooting,
**So that** I can increase log verbosity when diagnosing issues.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a log level selector dropdown in the Advanced tab with options: Verbose, Debug, Information, Warning, Error, Fatal
- [ ] AC-02: The currently active log level is pre-selected on page load with default "Information"
- [ ] AC-03: Log level selection must be one of the valid options (validated)
- [ ] AC-04: Settings Page displays the current log file path as a read-only text field
- [ ] AC-05: Settings Page provides an "Open Log Folder" button that opens the log directory in Windows Explorer
- [ ] AC-06: Saving persists the log level to `UserPreferences` table with key `logging.level`
- [ ] AC-07: Log level change takes effect immediately upon save without requiring application restart
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-017 | Settings Page SHALL provide a log level selector (Verbose, Debug, Information, Warning, Error, Fatal) with default "Information" | Should |
| FR-007-018 | Settings Page SHALL display the current log file path and provide an "Open Log Folder" button | Should |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-09-01 | Add Logging section to Advanced tab with log level dropdown, log path display, and Open Log Folder button | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-09-02 | Add SelectedLogLevel and LogFilePath properties to ViewModel | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-09-03 | Implement OpenLogFolderCommand that opens log directory in Windows Explorer using Process.Start | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-09-04 | Implement validation for log level (must be one of the six valid options) | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-09-05 | Persist log level to UserPreferences table on Save | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-09-06 | Apply log level change at runtime by reconfiguring the logging provider minimum level | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-09-07 | Load log level from UserPreferences table on page initialization with fallback to "Information" | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The Python application uses basic console logging. The C# migration uses the standard .NET `ILogger<T>` infrastructure with configurable log levels.
- The log level maps to .NET `LogLevel` enum: Verbose->Trace, Debug->Debug, Information->Information, Warning->Warning, Error->Error, Fatal->Critical.
- Validation rule: VR-007-012 (log level must be one of "Verbose", "Debug", "Information", "Warning", "Error", "Fatal").
- ConfigKey mapping: `logging.level` in `UserPreferences` table.
- The "Open Log Folder" button should use `Process.Start(new ProcessStartInfo { FileName = logFolderPath, UseShellExecute = true })` to open Windows Explorer.
- Log level change should take effect immediately by reconfiguring the `ILoggerFactory` minimum level, without requiring restart.
