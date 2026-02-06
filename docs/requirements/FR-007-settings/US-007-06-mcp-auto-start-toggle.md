# US-007-06: MCP Auto-Start Toggle

## User Story
**As a** System Admin,
**I want to** enable or disable MCP server auto-start,
**So that** I can control whether the AI assistant launches automatically.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a toggle/checkbox for MCP server auto-start in the MCP tab
- [ ] AC-02: The toggle defaults to off (disabled) for new installations
- [ ] AC-03: When enabled, the MCP SSE server starts automatically on application launch
- [ ] AC-04: When disabled, the MCP SSE server does not start on application launch (can still be started manually)
- [ ] AC-05: Saving persists the auto-start setting to `ServerConfigs` database table with key `mcp.auto_start`
- [ ] AC-06: The auto-start value is also saved to the `appsettings.json` MCP section
- [ ] AC-07: Auto-start setting is loaded from database on page initialization with fallback to false
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-011 | Settings Page SHALL provide a toggle for MCP server auto-start on application launch (default: off) | Must |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-06-01 | Add MCP auto-start toggle/checkbox to MCP tab in XAML bound to MCPAutoStart property | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-06-02 | Add MCPAutoStart boolean property to ViewModel with default value false | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-06-03 | Persist MCP auto-start to ServerConfigs table and appsettings.json on Save | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-06-04 | Load MCP auto-start from ServerConfigs table on page initialization with fallback to false | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-06-05 | Integrate auto-start check in application startup logic to conditionally launch MCP server | `ViewModels/SettingsViewModel.cs` | M |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The Python application uses the `--enable-mcp` CLI flag to control MCP auto-start. The C# migration replaces this with a persistent setting in the Settings Page.
- The persisted preference in `ServerConfigs` overrides the `appsettings.json` default, following the standard load order: JSON file (base) -> DB override.
- ConfigKey mapping: `mcp.auto_start` in `ServerConfigs` table with DataType "bool".
- The application startup logic should check this setting to determine whether to launch the MCP SSE server automatically.
