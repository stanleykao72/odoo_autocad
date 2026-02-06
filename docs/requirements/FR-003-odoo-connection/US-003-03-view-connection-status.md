# US-003-03: View Connection Status

## User Story
**As a** CAD Engineer,
**I want to** see the current Odoo connection status at all times,
**So that** I know whether data operations will succeed.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A connection status panel is always visible on the Odoo Integration Page
- [ ] AC-02: The status displays a colored indicator: green Ellipse for connected, red/gray Ellipse for disconnected
- [ ] AC-03: Descriptive text shows current state: "Connected", "Disconnected", or "Connecting..."
- [ ] AC-04: The `IsConnected` property returns true only when a valid session ID exists and authentication has succeeded
- [ ] AC-05: When connected, the panel displays the authenticated username and database name
- [ ] AC-06: When disconnected, the panel shows an appropriate disconnected state message
- [ ] AC-07: Status updates reactively when connection state changes (connect, disconnect, session expiry)
- [ ] AC-08: Other pages (Dashboard, BOQ, PR) can subscribe to connection state changes to update their enabled state

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-005 | The `IsConnected` property SHALL return true only when a valid session ID exists and authentication has succeeded | Must |
| FR-003-006 | Connection status SHALL be displayed as a colored indicator (green = connected, red/gray = disconnected) with descriptive text | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-03-01 | Create Connection Status Border panel with colored Ellipse indicators and text labels in XAML | `Views/Pages/OdooConnectionPage.xaml` | M |
| TASK-003-03-02 | Implement IsConnected property with proper change notification in ViewModel | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-03-03 | Implement ConnectionStatusText property ("Connected" / "Disconnected" / "Connecting...") | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-03-04 | Create BoolToColorConverter for green/red/gray status indicator binding | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-03-05 | Implement IOdooService.IsConnected property that checks for valid session ID | `OdooAutoCAD.Core/Odoo/OdooService.cs` | S |
| TASK-003-03-06 | Add connection state change event/callback so other ViewModels can subscribe to status changes | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | M |

## Dependencies
- Depends on: US-003-01
- Blocks: US-003-05, US-003-06, US-003-07

## Notes
- The status panel should use WPF `Border` with rounded corners and consistent theme from `App.xaml` resources.
- The `IsConnected` check is used by validation rules VR-003-008 and VR-003-009 to gate product sync and project search operations.
- Connection state events are important for cross-page reactivity: when the connection drops, BOQ and PR pages should disable their Odoo-dependent operations.
- The Python equivalent is `connected_odoo()` which returns True if `self.odoo` is not None.
