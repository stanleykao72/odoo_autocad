# US-003-04: Disconnect from Odoo

## User Story
**As a** CAD Engineer,
**I want to** disconnect from Odoo explicitly,
**So that** I can switch to a different server or database.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A "Disconnect" button is visible on the connection settings panel
- [ ] AC-02: Clicking "Disconnect" calls `DisconnectAsync()` and terminates the current Odoo session
- [ ] AC-03: After disconnection, session state is cleared (session cookies, uid, session ID)
- [ ] AC-04: The `IsConnected` property returns false after disconnection
- [ ] AC-05: The connection status indicator changes to red/gray with text "Disconnected"
- [ ] AC-06: Server info fields (version, database, username in status panel) are cleared
- [ ] AC-07: Product sync, project search, and sync-to-Odoo buttons are disabled after disconnection
- [ ] AC-08: The "Disconnect" button is disabled when already disconnected
- [ ] AC-09: Credential input fields remain populated so the user can reconnect or modify and reconnect

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-004 | The page SHALL provide a "Disconnect" button that terminates the current Odoo session and clears session state | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-04-01 | Add "Disconnect" button to the connection settings button bar in XAML, bound to DisconnectCommand | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-04-02 | Implement DisconnectCommand as IAsyncRelayCommand calling IOdooService.DisconnectAsync() | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-04-03 | Implement DisconnectAsync() in OdooService to clear session cookies, uid, and internal state | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-04-04 | Reset ViewModel status properties (IsConnected, ServerVersion, ConnectedDatabase, ConnectedUsername) on disconnect | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-04-05 | Bind Disconnect button IsEnabled to IsConnected so it is only active when connected | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-04-06 | Fire connection state change event to notify other pages of disconnection | `OdooAutoCAD.Core/Odoo/OdooService.cs` | S |

## Dependencies
- Depends on: US-003-01, US-003-03
- Blocks: None

## Notes
- Disconnection should be graceful: clear internal state without throwing exceptions even if the server is unreachable.
- The credential input fields (Server URL, Database, Username, Password) should remain populated after disconnect to facilitate quick reconnection or switching to a different server.
- Other pages that depend on Odoo connection (BOQ, PR) should reactively disable their operations when disconnect fires the connection state change event.
- The Python equivalent does not have an explicit disconnect; the C# version adds this for cleaner session management.
