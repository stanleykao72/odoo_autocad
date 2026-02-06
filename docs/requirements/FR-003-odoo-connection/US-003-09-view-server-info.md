# US-003-09: View Server Info

## User Story
**As a** Project Manager,
**I want to** view server information (version, database name) after connecting,
**So that** I can verify we are connected to the correct environment.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: On successful connection, a status panel displays the Odoo server version (e.g., "17.0")
- [ ] AC-02: The status panel displays the connected database name
- [ ] AC-03: The status panel displays the authenticated username
- [ ] AC-04: Server version is retrieved via `POST /web/webclient/version_info` after authentication
- [ ] AC-05: The server info panel is only populated when `IsConnected` is true
- [ ] AC-06: Server info fields are cleared when the user disconnects
- [ ] AC-07: The session status ("Active" / "Expired") is displayed alongside the connection info

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-007 | On successful connection, the page SHALL display server version, database name, and authenticated username in a status panel | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-09-01 | Add server info labels (Version, Database, Username, Session status) to the Connection Status panel in XAML | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-09-02 | Add ServerVersion, ConnectedDatabase, ConnectedUsername properties to ViewModel | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-09-03 | Implement GetStatusAsync() in OdooService to return OdooStatus record with version info | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-09-04 | Call `POST /web/webclient/version_info` after successful authentication to retrieve server version | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-09-05 | Populate server info properties in ViewModel after successful ConnectCommand execution | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-09-06 | Clear server info properties in ViewModel on DisconnectCommand execution | `ViewModels/OdooConnectionViewModel.cs` | S |

## Dependencies
- Depends on: US-003-01, US-003-03
- Blocks: None

## Notes
- The wireframe shows server info in the Connection Status panel: "Status: (o) Connected, Server Version: 17.0, Database: production, Username: admin, Last Sync: timestamp, Session: Active".
- The C# implementation retrieves server version via `POST /web/webclient/version_info`, which is a standard Odoo endpoint. The Python version did not use this endpoint.
- The `OdooStatus` record contains: IsConnected, ServerUrl, Database, Username, Version, ErrorMessage.
- This information is primarily for verification purposes -- ensuring the correct environment (production vs. staging) is connected.
