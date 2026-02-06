# TASKS: US-004-05 — ID Writeback

> **Parent US**: [US-004-05](US-004-05-id-writeback.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 2S + 2M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-04 (Push must succeed to get IDs for writeback — `SyncResult.Success` must be true)
- [ ] US-004-01 (MTextFormatter.UnFormat() must be available for product_no matching)

## Acceptance Criteria
- [ ] AC-01: After a successful push, the system automatically writes `header_id` to cell (row 0, column 8) for each legal table
- [ ] AC-02: For each data row (i > 1), the system writes the matched `detail_id` to cell (i, column 8) by matching `product_no` (after MText unformatting) to the Odoo response
- [ ] AC-03: Product number matching for writeback uses the `MTextFormatter.UnFormat()` equivalent to normalize cell text before comparison
- [ ] AC-04: All writeback COM operations are routed through `IGUIProxy.ExecuteInGuiAsync()` for STA thread safety
- [ ] AC-05: If writeback fails for a specific layout, an error message is shown for that layout: "Failed to write IDs back to AutoCAD table in layout '{layout_name}': {error}"
- [ ] AC-06: Already-pushed data remains in Odoo even if writeback fails; the user can retry writeback separately
- [ ] AC-07: The DataGrid DetailId column is updated to reflect the written-back values

---

## TASK-004-05-01: Implement set_layouts_tables_id equivalent via IGUIProxy for writing header_id and detail_id to tables

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | L |
| Depends On | TASK-004-05-02, TASK-004-05-03 |
| Blocks | TASK-004-05-05, TASK-004-05-06 |

### What to do
- Create a private method `ExecuteWritebackAsync(List<WritebackLayout> writebackData)` in `BOQManagerViewModel`
- Called automatically at the end of `ExecutePushAsync()` when push succeeds and writeback data is available
- For each layout in the writeback data, call `_guiProxy.ExecuteInGuiAsync("set_layouts_tables_id", parameters)` where parameters include:
  - `layout_name`: the layout name to identify which layout to write to
  - `header_id`: the value to write to cell (row 0, column 8)
  - `details`: a list of `(product_no, detail_id)` tuples for row matching
- The registered GUI proxy handler iterates each legal table in the layout:
  1. Write `header_id` to cell (row 0, column 8) via `IAutoCADService.SetTableValue(tableHandle, 0, 8, headerId)`
  2. For rows i > 1: read `product_no` from cell (i, 1), apply `MTextFormatter.UnFormat()`, find matching `detail_id` from the response, write to cell (i, 8)
  3. Row 1 (label row) is skipped
- If `product_no` match is not found for a row, leave that row's detail_id cell unchanged (matching Python behavior)
- Wrap each layout's writeback in try-catch for per-layout error handling

### How to verify
- [ ] header_id is written to cell (0, 8) for each legal table (AC-01)
- [ ] detail_id is written to cell (i, 8) for data rows (i > 1) by matching product_no (AC-02)
- [ ] Product matching uses MTextFormatter.UnFormat() for normalization (AC-03)
- [ ] All COM writes are routed through IGUIProxy (AC-04)

---

## TASK-004-05-02: Implement SetTableValue() in IAutoCADService for writing individual cell values in AutoCAD tables

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-05-01 |

### What to do
- Implement the `SetTableValue(string tableHandle, int row, int column, string value)` method in `AutoCADService`
- Use the AutoCAD COM API to locate the table entity by its handle
- Call the table COM object's `SetCellValue(row, column, value)` or `SetText(row, column, value)` method
- Handle COM exceptions (e.g., invalid handle, out-of-range row/column) and throw descriptive exceptions
- Validate that `row >= 0` and `column >= 0` before calling COM
- Implement `GetTableValue(string tableHandle, int row, int column)` as well for reading cell values during the matching process
- Add logging for each write operation: `_logger.LogDebug("Writing '{Value}' to table {Handle} cell ({Row}, {Col})", value, tableHandle, row, column)`

### How to verify
- [ ] SetTableValue writes the specified value to the correct cell in an AutoCAD table (AC-01, AC-02)
- [ ] GetTableValue reads cell values for product_no matching (AC-02)

---

## TASK-004-05-03: Register writeback handler in IGUIProxy for STA thread execution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-05-01 |

### What to do
- Register a `"set_layouts_tables_id"` handler with `IGUIProxy.RegisterHandler()` during application startup
- The handler receives parameters: `layout_name` (string), `header_id` (string), `details` (list of product_no/detail_id pairs)
- The handler implementation:
  1. Switches to the specified layout via `IAutoCADService.SwitchToLayout(layoutName)`
  2. Gets all tables via `IAutoCADService.GetTables()` and filters for legal tables using `BOQProcessor.IsLegalTable()`
  3. For each legal table: writes `header_id` to cell (0, 8) and iterates rows 2+ to match and write `detail_id`
  4. Returns a success/failure result
- Handle the case where the layout is not found or has no legal tables
- Ensure the handler executes synchronously on the STA thread (wrap result in `Task.FromResult()`)

### How to verify
- [ ] `ExecuteInGuiAsync("set_layouts_tables_id")` executes writeback on the GUI thread (AC-04)
- [ ] Per-layout writeback succeeds when layout exists with legal tables (AC-01, AC-02)

---

## TASK-004-05-04: Implement get_detail_id_index equivalent: match product_no to detail_id from Odoo response

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-05-01 |

### What to do
- Add a static method `public static string? FindDetailId(IEnumerable<(string ProductNo, string DetailId)> details, string productNo)` to `BOQProcessor`
- The method applies `MTextFormatter.UnFormat()` to the input `productNo` for normalization
- Iterates the `details` list and returns the `DetailId` for the first entry where `ProductNo` matches (case-insensitive, trimmed)
- Returns `null` if no match is found (the caller leaves the cell unchanged in this case)
- This is the C# equivalent of the Python `get_detail_id_index(detail_list, product_no)` function

### How to verify
- [ ] ProductNo matching is case-insensitive and uses MText unformatting (AC-03)
- [ ] Returns null when no match is found, leaving the cell unchanged (AC-02)

---

## TASK-004-05-05: Update BOQEntryRow.DetailId in the DataGrid after successful writeback

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-05-01 |
| Blocks | None |

### What to do
- After each layout's writeback completes successfully in `ExecuteWritebackAsync()`:
  - Find the corresponding `BOQEntryRow` items in `AllEntries` by `LayoutName`
  - Update each entry's `DetailId` property with the written-back `detail_id` value
  - Update the parent `BOQLayoutGroup.HeaderId` with the written-back `header_id`
- Call `OnPropertyChanged` on each modified entry to trigger DataGrid cell refresh
- This provides immediate visual feedback that the writeback completed for each row

### How to verify
- [ ] DataGrid DetailId column shows the written-back values after writeback (AC-07)
- [ ] HeaderId on the layout group is updated (AC-07)

---

## TASK-004-05-06: Handle per-layout writeback errors with user-facing messages and retry option

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-05-01 |
| Blocks | None |

### What to do
- In `ExecuteWritebackAsync()`, wrap each layout's writeback in a try-catch
- On failure, add an error message to a writeback errors list: `$"Failed to write IDs back to AutoCAD table in layout '{layoutName}': {ex.Message}"`
- Display writeback errors in the push result panel (append to `LastPushResult` or use a dedicated `WritebackErrors` collection)
- Ensure that a writeback failure for one layout does not prevent writeback attempts for other layouts
- Emphasize in the UI that Odoo data is preserved even when writeback fails: "Data has been saved to Odoo. Writeback to AutoCAD failed for the following layouts:"
- Add a "Retry Writeback" `IAsyncRelayCommand` that re-attempts writeback for layouts that previously failed
- Log each writeback error via `_logger.LogError()`

### How to verify
- [ ] Per-layout writeback errors show the specific layout name and error message (AC-05)
- [ ] Pushed data remains in Odoo when writeback fails (AC-06)
- [ ] User can retry writeback separately (AC-06)

---

## Dependency Graph
```
TASK-004-05-02 (SetTableValue Service) ────┐
                                            │
TASK-004-05-03 (IGUIProxy Writeback Handler)┤
                                            │
TASK-004-05-04 (FindDetailId Matching) ─────┤
                                            ▼
                                  TASK-004-05-01 (Writeback Orchestration)
                                            │
                                            ├──▶ TASK-004-05-05 (Update DataGrid)
                                            └──▶ TASK-004-05-06 (Error Handling)

Tasks 02, 03, 04 can begin in parallel (no interdependencies).
Task 01 depends on all three (02, 03, 04).
Tasks 05 and 06 depend on 01.
```
