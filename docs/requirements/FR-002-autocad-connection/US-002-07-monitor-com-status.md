# US-002-07: Monitor COM Status

## User Story
**As a** System Admin,
**I want to** monitor COM connection status,
**So that** I can troubleshoot connectivity issues.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: The page displays a persistent connection status indicator showing "Connected" or "Disconnected" with a visual cue (e.g., green/red icon)
- [ ] AC-02: When connected, AutoCAD version and application info are displayed
- [ ] AC-03: All COM operations are routed through IGUIProxy to ensure STA thread safety
- [ ] AC-04: If a COM operation times out, the error message "AutoCAD operation timed out. The application may be busy." is displayed with a retry option
- [ ] AC-05: If a COM thread conflict occurs, the system transparently retries via the GUI proxy
- [ ] AC-06: Connection status is continuously monitored; if AutoCAD disconnects unexpectedly, the status updates to "Disconnected" and dependent controls are disabled
- [ ] AC-07: A "Refresh Status" button allows manual re-check of the COM connection health

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-003 | Page SHALL display connection status (Connected/Disconnected) with visual indicator | Must |
| FR-002-004 | Page SHALL display AutoCAD version and application info when connected | Should |
| FR-002-007 | All COM operations SHALL execute on the GUI/STA thread via IGUIProxy | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-07-01 | Add connection status indicator (icon + text) and AutoCAD version info display to XAML | `Views/Pages/AutoCADPage.xaml` | S |
| TASK-002-07-02 | Add "Refresh Status" button to the Connection Status panel | `Views/Pages/AutoCADPage.xaml` | S |
| TASK-002-07-03 | Implement GetStatusAsync() in AutoCADService to check COM connection health and retrieve version info | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-07-04 | Define GetStatusAsync() returning AutoCADStatus in IAutoCADService interface | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | S |
| TASK-002-07-05 | Implement RefreshStatusCommand in ViewModel to invoke status check | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-07-06 | Implement periodic connection health check using DispatcherTimer (100ms interval for IGUIProxy.ProcessRequests) | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-002-07-07 | Implement COM timeout detection and retry logic via IGUIProxy | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | M |
| TASK-002-07-08 | Handle unexpected disconnection: update IsConnected, disable dependent controls, show notification | `ViewModels/AutoCADViewModel.cs` | M |

## Dependencies
- Depends on: None (standalone -- can be developed independently)
- Blocks: None

## Notes
- The WPF DispatcherTimer running at 100ms intervals calls `IGUIProxy.ProcessRequests()` to process queued COM operations. This same mechanism can be used to periodically check connection health.
- COM thread safety is critical: the MCP server runs on an MTA thread while AutoCAD COM requires STA. The IGUIProxy bridges this gap by queuing operations for execution on the WPF Dispatcher thread.
- The Python implementation monitors connection state and displays it via the GUI. The C# implementation should use the MVVM pattern with observable properties for status binding.
- Error scenarios to handle: AutoCAD not running, COM connection failed, no active document, ActiveDocument retry exhausted, and COM thread conflict (timeout).
- The AutoCADStatus model should include: IsConnected (bool), Version (string), ApplicationInfo (string), LastChecked (DateTime).
