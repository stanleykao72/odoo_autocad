# US-001-01: View Connection Status

## User Story
**As a** CAD Engineer,
**I want to** see connection status for AutoCAD and Odoo at a glance,
**So that** I know if I can start working immediately.

## Parent Feature
- **FR**: [FR-001-dashboard](../FR-001-dashboard/FR-001-dashboard.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Dashboard displays an AutoCAD connection status card with icon, label, and connected/disconnected state
- [ ] AC-02: Dashboard displays an Odoo connection status card with icon, label, and connected/disconnected state
- [ ] AC-03: Connected state displays with a green indicator; disconnected state displays with a red/gray indicator
- [ ] AC-04: Connection status cards update in real-time when the connection state changes (no manual refresh required)
- [ ] AC-05: AutoCAD status card shows the current document name when connected
- [ ] AC-06: Odoo status card shows the server URL when connected
- [ ] AC-07: When AutoCAD connection fails, the status card displays "Unable to connect to AutoCAD. Please ensure AutoCAD is running."
- [ ] AC-08: When Odoo connection fails, the status card displays "Unable to connect to Odoo server. Check connection settings."

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-001-001 | Dashboard SHALL display AutoCAD connection status as a card with icon, label, and connected/disconnected state | Must |
| FR-001-002 | Dashboard SHALL display Odoo connection status as a card with icon, label, and connected/disconnected state | Must |
| FR-001-004 | Connection status cards SHALL update in real-time when connection state changes | Must |
| FR-001-005 | Connected state SHALL display with green indicator; disconnected with red/gray indicator | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-001-01-01 | Create DashboardPage XAML layout with connection status cards (AutoCAD, Odoo) using Border with rounded corners, icon, title, and status indicator | `Views/Pages/DashboardPage.xaml` | M |
| TASK-001-01-02 | Add DashboardViewModel properties: IsAutoCADConnected, AutoCADStatus, AutoCADDocumentName, IsOdooConnected, OdooStatus, OdooServerUrl | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-01-03 | Implement BoolToColorConverter for green (connected) / red-gray (disconnected) status indicator binding | `Views/Pages/DashboardPage.xaml` | S |
| TASK-001-01-04 | Subscribe DashboardViewModel to IAutoCADService and IOdooService status change events for real-time updates | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-01-05 | Register DashboardViewModel as singleton in DI container to maintain state across navigation | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-01-06 | Implement error message display in status cards when connection attempts fail | `ViewModels/DashboardViewModel.cs` | S |

## Dependencies
- Depends on: FR-008 (UI Framework - WPF/MVVM infrastructure and CommunityToolkit.Mvvm must be in place)
- Depends on: IAutoCADService and IOdooService interfaces being defined
- Blocks: US-001-02 (quick connect buttons rely on connection status display)

## Notes
- Python uses emoji indicators (green circle/red circle) for status; the C# implementation should use WPF Ellipse elements with Fill color bound via BoolToColorConverter.
- Python displays status in the top bar; the C# dashboard consolidates status into dedicated cards on the dashboard page.
- The ViewModel should use `ObservableObject` from CommunityToolkit.Mvvm to enable property change notifications.
- Connection status event subscriptions should be properly disposed when the ViewModel is deactivated to prevent memory leaks.
