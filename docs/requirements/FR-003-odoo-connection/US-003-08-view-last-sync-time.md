# US-003-08: View Last Sync Time

## User Story
**As a** CAD Engineer,
**I want to** see the last successful sync timestamp,
**So that** I know how fresh my local product data is.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The last synchronization timestamp is displayed on the Odoo Integration Page
- [ ] AC-02: The timestamp is retrieved via `GetLastSyncTime()` from the IOdooService
- [ ] AC-03: The timestamp is formatted in the user's local timezone
- [ ] AC-04: The timestamp updates automatically after each successful sync operation
- [ ] AC-05: If no sync has ever been performed, the display shows "Never synced" or equivalent
- [ ] AC-06: The sync timestamp is visible in both the Connection Status panel and the Sync Controls panel

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-021 | The page SHALL display the last synchronization timestamp via `GetLastSyncTime()`, formatted in the user's local timezone | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-08-01 | Add last sync timestamp display label in the Connection Status panel and Sync Controls panel | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-08-02 | Add LastSyncTime nullable DateTime property to ViewModel with formatted display binding | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-08-03 | Implement GetLastSyncTime() in OdooService to return the stored last sync timestamp | `OdooAutoCAD.Core/Odoo/OdooService.cs` | S |
| TASK-003-08-04 | Persist last sync time to local database so it survives application restarts | `OdooAutoCAD.Data/Context/AppDbContext.cs` | M |
| TASK-003-08-05 | Update LastSyncTime after each successful sync operation (products, BOQ, parameters) | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-08-06 | Create DateTime-to-string converter for local timezone formatting with "Never synced" fallback | `Views/Pages/OdooConnectionPage.xaml` | S |

## Dependencies
- Depends on: US-003-05
- Blocks: None

## Notes
- The timestamp should be formatted using the user's local timezone settings, e.g., "2026-02-06 14:30" as shown in the wireframe.
- The last sync time should be persisted locally (in SQLite) so it is available across application restarts.
- Each sync type (Products, BOQ, Parameters) may have its own last sync timestamp, but this US focuses on the most recent overall sync time.
- The Sync Controls panel in the wireframe shows "Last Sync: 2026-02-06 14:30" alongside the sync action button.
