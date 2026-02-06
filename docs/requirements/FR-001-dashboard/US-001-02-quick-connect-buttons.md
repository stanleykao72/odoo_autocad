# US-001-02: Quick Connect Buttons

## User Story
**As a** CAD Engineer,
**I want to** have quick-connect buttons on the dashboard,
**So that** I can establish connections without navigating to settings.

## Parent Feature
- **FR**: [FR-001-dashboard](../FR-001-dashboard/FR-001-dashboard.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Dashboard provides a "Connect to AutoCAD" quick-action button within the AutoCAD status card
- [ ] AC-02: Dashboard provides a "Connect to Odoo" quick-action button within the Odoo status card
- [ ] AC-03: Quick-connect buttons are disabled and show "Connected" state when the respective service is already connected
- [ ] AC-04: Clicking "Connect to AutoCAD" initiates an async connection attempt to AutoCAD
- [ ] AC-05: Clicking "Connect to Odoo" initiates an async connection attempt to Odoo
- [ ] AC-06: Quick-connect buttons check current connection state before attempting connection (VR-001-001)
- [ ] AC-07: Dashboard provides shortcut buttons for common workflows: "Get Parameters", "Push to BOQ", and "Transfer BOQ to PR"
- [ ] AC-08: "Get Parameters" shortcut is only enabled when both AutoCAD AND Odoo are connected (VR-001-003)
- [ ] AC-09: "Push to BOQ" shortcut is only enabled when Odoo is connected and project context is available (VR-001-004)
- [ ] AC-10: Shortcut buttons navigate to the corresponding application page when clicked

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-001-006 | Dashboard SHALL provide a "Connect to AutoCAD" quick-action button | Must |
| FR-001-007 | Dashboard SHALL provide a "Connect to Odoo" quick-action button | Must |
| FR-001-008 | Quick-action buttons SHALL be disabled and show "Connected" state when already connected | Should |
| FR-001-009 | Dashboard SHOULD provide shortcuts to common workflows (Get Parameters, Push to BOQ) | Could |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-001-02-01 | Add "Connect to AutoCAD" and "Connect to Odoo" buttons inside their respective status cards in XAML, bound to commands | `Views/Pages/DashboardPage.xaml` | S |
| TASK-001-02-02 | Implement ConnectAutoCADCommand using RelayCommand with async execution calling IAutoCADService.ConnectAsync() | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-02-03 | Implement ConnectOdooCommand using RelayCommand with async execution calling IOdooService.ConnectAsync() | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-02-04 | Add CanExecute logic to connect commands: disable when already connected, enable when disconnected | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-02-05 | Create Quick Actions section in XAML with shortcut buttons for Get Parameters, Push to BOQ, and Transfer BOQ to PR | `Views/Pages/DashboardPage.xaml` | M |
| TASK-001-02-06 | Implement NavigateToParametersCommand, NavigateToBOQCommand, and NavigateToPRCommand using INavigationService.NavigateTo() | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-02-07 | Add CanExecute validation: Get Parameters requires AutoCAD AND Odoo connected; Push to BOQ requires Odoo connected and project context | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-02-08 | Re-evaluate command CanExecute states when connection status changes (call NotifyCanExecuteChanged on related commands) | `ViewModels/DashboardViewModel.cs` | S |

## Dependencies
- Depends on: US-001-01 (connection status cards must exist before adding buttons to them)
- Depends on: IAutoCADService.ConnectAsync() and IOdooService.ConnectAsync() being implemented
- Depends on: INavigationService for workflow shortcut navigation
- Blocks: none

## Notes
- The connect commands should be async (using `AsyncRelayCommand` from CommunityToolkit.Mvvm) to avoid blocking the UI thread during connection attempts.
- When connection state changes, all related commands must call `NotifyCanExecuteChanged()` to update button enabled/disabled states automatically.
- Validation rules VR-001-001 through VR-001-004 from the FR document must be enforced through the CanExecute logic of each command.
- The Quick Actions section (FR-001-009) is priority "Could", so it can be deferred if needed, but the connect buttons (FR-001-006, FR-001-007) are priority "Must".
- Python uses 150x40px buttons; the C# implementation should use WPF styled buttons that match the overall application theme.
