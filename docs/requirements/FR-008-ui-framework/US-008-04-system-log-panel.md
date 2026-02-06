# US-008-04: System Log Panel

## User Story
**As a** CAD Engineer,
**I want to** view system log messages at the bottom of the screen,
**So that** I can monitor operations and diagnose issues in real-time.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A status bar at the bottom of the content area displays a `StatusMessage` text, the current date/time, and the application version.
- [ ] AC-02: A dedicated system log panel is available (as a collapsible bottom panel or a separate toggle overlay) to display operational log messages.
- [ ] AC-03: The system log panel auto-scrolls to the latest entry when new messages are added.
- [ ] AC-04: Log entries include timestamps in the format `[YYYY-MM-DD HH:mm:ss]` and category prefixes (e.g., "[SSE GUI]", "[GUI Proxy]", "[Odoo]").
- [ ] AC-05: The status bar displays three columns: status message (left, fills remaining space), date/time (center-right), and version string (far right).
- [ ] AC-06: The status bar uses `SurfaceBrush` background with a top border separator.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-027 | Status bar with StatusMessage, date/time, and version | Must |
| FR-008-028 | Dedicated system log panel for operational log messages | Should |
| FR-008-029 | System log panel auto-scroll to latest entry | Should |
| FR-008-030 | Log entries with timestamps and category prefixes | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-04-01 | Define status bar XAML with three-column layout (StatusMessage, DateTime, Version) | `Views/MainWindow.xaml` | S |
| TASK-008-04-02 | Add StatusMessage, CurrentDateTime, and version observable properties to MainViewModel | `ViewModels/MainViewModel.cs` | S |
| TASK-008-04-03 | Implement a collapsible log panel XAML below the status bar or as a toggle overlay with ScrollViewer | `Views/MainWindow.xaml` | M |
| TASK-008-04-04 | Create a Serilog sink or log collection that feeds log entries to the UI panel with timestamps and category prefixes | `ViewModels/MainViewModel.cs` | M |
| TASK-008-04-05 | Implement auto-scroll behavior on the log panel ScrollViewer when new entries are added | `Views/MainWindow.xaml.cs` | S |
| TASK-008-04-06 | Add a DispatcherTimer to update CurrentDateTime every second | `ViewModels/MainViewModel.cs` | S |

## Dependencies
- Depends on: US-008-01 (main window layout must include the content area where the status bar and log panel reside)
- Blocks: None

## Notes
- The Python implementation uses a "系統日誌" (System Log) heading with a `ScrolledText` widget (black background, white text) in the bottom area spanning both columns. The C# port may use a cleaner approach with a collapsible `Expander` or a toggle button.
- Serilog is configured in `App.OnStartup` and can be extended with a custom sink that writes to an `ObservableCollection<string>` bound to the log panel.
- The status bar is always visible; the detailed log panel may be hidden by default and toggled on demand.
- Log categories from the Python implementation include: "[SSE GUI]", "[GUI Proxy]", "[Odoo]", "[AutoCAD]".
