# US-004-10: Clear IDs

## User Story
**As a** CAD Engineer,
**I want to** clear existing IDs from all tables before a fresh push,
**So that** I can reset and re-push without stale references.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Clicking "Clear All IDs" clears column 8 across all layouts and tables in the active drawing
- [ ] AC-02: The clear operation iterates all non-Model layouts and all legal tables within each layout
- [ ] AC-03: For each table, column 8 is cleared for all rows except row 1 (the label row)
- [ ] AC-04: The operation is executed via `IGUIProxy.ExecuteInGuiAsync()` for STA thread safety
- [ ] AC-05: After clearing, the DataGrid DetailId column is updated to reflect empty values
- [ ] AC-06: A confirmation dialog is shown before clearing: "This will clear all Odoo IDs from the drawing. Continue?"
- [ ] AC-07: A success message is shown after clearing completes

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-028 | Provide "Clear All IDs" button that clears column 8 across all layouts and tables (equivalent to clear_all_tables_id()) | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-10-01 | Add "Clear All IDs" button to extraction panel in BOQ Manager page | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-10-02 | Implement ClearAllIdsCommand that calls IGUIProxy for COM clear operation | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-10-03 | Implement clear_all_tables_id equivalent in IAutoCADService: iterate layouts, iterate legal tables, clear column 8 | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | M |
| TASK-004-10-04 | Register clear handler in IGUIProxy for STA thread execution | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | S |
| TASK-004-10-05 | Add confirmation dialog before executing clear operation | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-10-06 | Update DataGrid DetailId values to empty after clear completes | `ViewModels/BOQManagerViewModel.cs` | S |

## Dependencies
- Depends on: US-002-01 (AutoCAD connected)
- Blocks: None

## Notes
- This operation is equivalent to the Python `clear_all_tables_id()` which calls `clear_table_id(layout)` for each non-Model layout.
- The Python `clear_table_id(layout)` clears column 8 for all rows except row 1 in each legal table within the layout.
- Row 0 contains `header_id` and rows 2+ contain `detail_id` values -- both are cleared. Row 1 is the column label row and is preserved.
- This is a destructive operation on the AutoCAD drawing, so a confirmation dialog is required before execution.
- The clear operation must be routed through `IGUIProxy` for the same COM thread safety reasons as extraction and writeback.
- After clearing, if the DataGrid already has data loaded, the `DetailId` property on each `BOQEntryRow` should be set to empty string.
