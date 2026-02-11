# US-008-04: System Log Panel

## User Story
**As a** CAD Engineer,
**I want to** view system log messages at the bottom of the screen,
**So that** I can monitor operations and diagnose issues in real-time.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [x] AC-01: A status bar at the bottom of the content area displays a `StatusMessage` text, the current date/time, and the application version.
- [x] AC-02: A dedicated system log panel is available (as a collapsible bottom panel or a separate toggle overlay) to display operational log messages.
- [x] AC-03: The system log panel auto-scrolls to the latest entry when new messages are added.
- [x] AC-04: Log entries include timestamps in the format `HH:mm:ss.fff` and source prefixes (e.g., "Odoo", "AutoCAD", "Settings", "App").
- [x] AC-05: The status bar displays three columns: status message (left, fills remaining space), date/time (center-right), and version string (far right).
- [x] AC-06: The status bar uses `SurfaceBrush` background with a top border separator.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-027 | Status bar with StatusMessage, date/time, and version | Must |
| FR-008-028 | Dedicated system log panel for operational log messages | Should |
| FR-008-029 | System log panel auto-scroll to latest entry | Should |
| FR-008-030 | Log entries with timestamps and category prefixes | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate | Status |
|---------|-------------|-------------|----------|--------|
| TASK-008-04-01 | Define status bar XAML with three-column layout (StatusMessage, DateTime, Version) | `Views/MainWindow.xaml` | S | Done (Sprint 1) |
| TASK-008-04-02 | Add StatusMessage, CurrentDateTime, and version observable properties to MainViewModel | `ViewModels/MainViewModel.cs` | S | Done (Sprint 1) |
| TASK-008-04-03 | Create IAppLogService/AppLogService singleton + collapsible LogPanelControl in MainWindow | `Services/IAppLogService.cs`, `Services/AppLogService.cs`, `Views/Controls/LogPanelControl.xaml/.cs`, `Views/MainWindow.xaml` | M | Done |
| TASK-008-04-04 | Create LogsPage full-page viewer + "Logs" nav button | `Views/Pages/LogsPage.xaml/.cs`, `Views/MainWindow.xaml` | M | Done |
| TASK-008-04-05 | Implement auto-scroll + source/level filtering in LogPanelControl | `Views/Controls/LogPanelControl.xaml.cs` | S | Done |
| TASK-008-04-06 | Add DispatcherTimer to update CurrentDateTime every second | `ViewModels/MainViewModel.cs` | S | Done (Sprint 1) |
| TASK-008-04-07 | Migrate OdooConnectionViewModel to IAppLogService, remove per-page log | `ViewModels/OdooConnectionViewModel.cs`, `Views/Pages/OdooPage.xaml` | M | Done |
| TASK-008-04-08 | Add IAppLogService log calls to AutoCADViewModel and SettingsViewModel | `ViewModels/AutoCADViewModel.cs`, `ViewModels/SettingsViewModel.cs` | S | Done |

## Dependencies
- Depends on: US-008-01 (main window layout must include the content area where the status bar and log panel reside)
- Blocks: None

## Actual Implementation Notes
The implementation deviated from the original spec in a beneficial way:

1. **Centralized `IAppLogService`** instead of a Serilog sink to `MainViewModel.LogEntries`. This provides a proper service that any ViewModel can inject and log to, with thread-safe dispatcher marshaling, Serilog forwarding, and a 2000-entry cap.

2. **Reusable `LogPanelControl` UserControl** instead of inline XAML in MainWindow. This control is used in both the collapsible bottom panel and the full-page LogsPage.

3. **Dedicated `LogsPage`** accessible via "Logs" nav button for full-page browsing with the same filtering capabilities.

4. **Source/Level filtering** via ComboBoxes with `CollectionViewSource`, exceeding the original AC requirements.

5. **Color-coded log levels**: Gray=Debug, Black=Info, Orange=Warning, Red=Error.

6. **Per-ViewModel migration**: `OdooConnectionViewModel` private `LogEntries`/`AddLog()` replaced with shared `IAppLogService.Log()`. `AutoCADViewModel` and `SettingsViewModel` also integrated.

### Key Files Created
- `src/OdooAutoCAD.App/Services/IAppLogService.cs` — `AppLogLevel` enum, `AppLogEntry` record, `IAppLogService` interface
- `src/OdooAutoCAD.App/Services/AppLogService.cs` — Singleton, thread-safe, 2000 cap, Serilog forwarding
- `src/OdooAutoCAD.App/Views/Controls/LogPanelControl.xaml/.cs` — Reusable UserControl with filtering
- `src/OdooAutoCAD.App/Views/Pages/LogsPage.xaml/.cs` — Full-page log viewer

### Key Files Modified
- `App.xaml.cs` — `IAppLogService` registered as singleton in DI
- `MainWindow.xaml/.cs` — Collapsible Expander with LogPanelControl, "Logs" nav button
- `OdooConnectionViewModel.cs` — Migrated from private log to shared `IAppLogService`
- `OdooPage.xaml` — Removed per-page "Connection Log" panel
- `AutoCADViewModel.cs` — Added `IAppLogService` log calls
- `SettingsViewModel.cs` — Added `IAppLogService` log calls
- `SettingsViewModelTests.cs` — Updated for new constructor parameter
