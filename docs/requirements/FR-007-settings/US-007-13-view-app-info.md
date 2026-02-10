# US-007-13: View App Info

## User Story
**As a** CAD Engineer,
**I want to** see the current application version and database path,
**So that** I can report environment details when seeking support.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [x] AC-01: Settings Page displays the application version number (e.g., "6.0.0") as a read-only field in the Advanced tab under Application Info section
- [x] AC-02: Settings Page displays the database file path (e.g., `C:\Users\...\database.db`) as a read-only field
- [x] AC-03: Settings Page displays the config file path (e.g., `appsettings.json`) as a read-only field
- [x] AC-04: Settings Page displays the .NET Runtime version (e.g., ".NET 8.0") as a read-only field
- [x] AC-05: All Application Info fields are read-only and cannot be edited by the user
- [x] AC-06: Application info values are populated automatically on page load from runtime metadata
- [x] AC-07: Version number is sourced from the assembly version or application metadata
- [x] AC-08: Database path is sourced from the EF Core connection string or AppDbContext configuration

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-025 | Settings Page SHALL display read-only application info: Version, Database Path, Config File Path, .NET Runtime version | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-13-01 | Add Application Info section to Advanced tab with read-only fields for Version, Database Path, Config File Path, .NET Runtime | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-13-02 | Add AppVersion, DatabasePath, ConfigFilePath, and DotNetRuntime read-only properties to ViewModel | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-13-03 | Populate AppVersion from Assembly.GetExecutingAssembly().GetName().Version or application metadata | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-13-04 | Populate DatabasePath from AppDbContext connection string or EF Core configuration | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-13-05 | Populate ConfigFilePath from ConfigurationLoader or AppDomain.CurrentDomain.BaseDirectory | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-13-06 | Populate DotNetRuntime from System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-13-07 | Style Application Info fields as read-only with distinct visual treatment (e.g., gray background, no border) | `Views/Pages/SettingsPage.xaml` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- All fields in this section are read-only. They exist to help users and support staff quickly identify the runtime environment when troubleshooting.
- Version number can be retrieved using `Assembly.GetExecutingAssembly().GetName().Version` or from a dedicated version constant/resource.
- Database path can be extracted from the SQLite connection string configured in `AppDbContext`, typically `Data Source=...`.
- .NET Runtime version is available via `System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription` (e.g., ".NET 8.0.0").
- Config file path should show the resolved absolute path to `appsettings.json` using `AppDomain.CurrentDomain.BaseDirectory` or the `IConfiguration` root path.
- Consider making the database path and config file path copyable (e.g., with a small copy-to-clipboard button) for easier sharing in support tickets.
