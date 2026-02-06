# US-007-10: Clear Cache

## User Story
**As a** System Admin,
**I want to** clear the local cache and synchronization data,
**So that** I can reset the application state when data becomes stale.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a "Clear Cache" button in the Advanced tab under Data Management section
- [ ] AC-02: Clicking "Clear Cache" shows a confirmation dialog with message: "This will remove {n} cached BOQ entries and {m} product mappings. This action cannot be undone. Continue?"
- [ ] AC-03: The confirmation dialog displays the actual record counts for BOQCache and ProductMapping tables
- [ ] AC-04: Confirming the dialog removes all records from `BOQCache` and `ProductMapping` database tables
- [ ] AC-05: Confirming the dialog also resets relevant `SyncLog` entries
- [ ] AC-06: Confirming the dialog invalidates any in-memory caches held by `IOdooService`
- [ ] AC-07: After successful cache clear, a success message with deleted record counts is displayed
- [ ] AC-08: Cancelling the confirmation dialog takes no action and returns to the settings page
- [ ] AC-09: If cache clear fails, an error dialog is shown: "Failed to clear cache. Error: {details}"
- [ ] AC-10: Cache clear operation is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-019 | Settings Page SHALL provide a "Clear Cache" button that removes all BOQCache and ProductMapping records from the local database | Must |
| FR-007-020 | Clear Cache SHALL require confirmation dialog before executing and display affected record counts | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-10-01 | Add "Clear Cache" button to Data Management section of Advanced tab in XAML | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-10-02 | Implement ClearCacheCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-10-03 | Query BOQCache and ProductMapping record counts for confirmation dialog | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-10-04 | Implement confirmation dialog with record counts and Cancel/Continue buttons | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-10-05 | Execute RemoveRange on BOQCache and ProductMapping tables via AppDbContext | `OdooAutoCAD.Data/Context/AppDbContext.cs` | M |
| TASK-007-10-06 | Invalidate in-memory caches in IOdooService after database clear | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |
| TASK-007-10-07 | Add success/error result feedback after cache clear operation | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-10-08 | Log cache clear operation to SyncLog table | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The "Clear Cache" scope includes: all rows from `BOQCache` table, all rows from `ProductMapping` table, reset of relevant `SyncLog` entries, and invalidation of any in-memory caches held by `IOdooService`.
- The confirmation dialog must show actual record counts, not estimates. This requires querying the database before showing the dialog.
- Error handling should catch `DbUpdateException` and other EF Core exceptions, displaying the inner exception message to the user.
- This operation is irreversible. The confirmation dialog should use warning-level styling (e.g., yellow/orange icon) to emphasize the destructive nature.
