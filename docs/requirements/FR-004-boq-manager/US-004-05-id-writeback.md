# US-004-05: ID Writeback

## User Story
**As a** CAD Engineer,
**I want to** have returned Odoo IDs written back to AutoCAD tables,
**So that** my drawings reflect the Odoo record references for traceability.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: After a successful push, the system automatically writes `header_id` to cell (row 0, column 8) for each legal table
- [ ] AC-02: For each data row (i > 1), the system writes the matched `detail_id` to cell (i, column 8) by matching `product_no` (after MText unformatting) to the Odoo response
- [ ] AC-03: Product number matching for writeback uses the `MTextFormatter.UnFormat()` equivalent to normalize cell text before comparison
- [ ] AC-04: All writeback COM operations are routed through `IGUIProxy.ExecuteInGuiAsync()` for STA thread safety
- [ ] AC-05: If writeback fails for a specific layout, an error message is shown for that layout: "Failed to write IDs back to AutoCAD table in layout '{layout_name}': {error}"
- [ ] AC-06: Already-pushed data remains in Odoo even if writeback fails; the user can retry writeback separately
- [ ] AC-07: The DataGrid DetailId column is updated to reflect the written-back values

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-020 | Write returned header_id and detail_id values back into AutoCAD table cells (column 8) via IAutoCADService.SetTableValue(), matching detail_id to rows by product_no | Must |
| FR-004-029 | Set header_id at cell (0, 8) and detail_id at cell (i, 8) for data rows (i > 1), matching by product_no using MText unformatting | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-05-01 | Implement set_layouts_tables_id equivalent via IGUIProxy for writing header_id and detail_id to tables | `ViewModels/BOQManagerViewModel.cs` | L |
| TASK-004-05-02 | Implement SetTableValue() in IAutoCADService for writing individual cell values in AutoCAD tables | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | M |
| TASK-004-05-03 | Register writeback handler in IGUIProxy for STA thread execution | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | M |
| TASK-004-05-04 | Implement get_detail_id_index equivalent: match product_no (MText-unformatted) to detail_id from Odoo response | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-05-05 | Update BOQEntryRow.DetailId in the DataGrid after successful writeback | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-05-06 | Handle per-layout writeback errors with user-facing messages and retry option | `ViewModels/BOQManagerViewModel.cs` | M |

## Dependencies
- Depends on: US-004-04 (push must succeed to get IDs for writeback)
- Blocks: None

## Notes
- This is the third step of the Python three-step workflow (`autocad_util.set_layouts_tables_id(boq_list)`). In C#, it executes automatically after a successful push.
- The writeback must use `IGUIProxy.ExecuteInGuiAsync("set_layouts_tables_id", parameters)` to ensure COM thread safety.
- Row matching logic: for each layout table, iterate data rows (i > 1), read `product_no` from cell (i, 1), apply MText unformatting, and find the corresponding `detail_id` from the Odoo response list using the same matching as the Python `get_detail_id_index()`.
- Row 1 is typically a header/label row and is skipped during writeback (only row 0 gets `header_id`, and rows 2+ get `detail_id`).
- If a `product_no` match is not found for a row, that row's `detail_id` cell is left unchanged (same behavior as Python reference).
