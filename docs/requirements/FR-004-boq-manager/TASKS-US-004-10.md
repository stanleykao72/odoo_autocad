# TASKS: US-004-10 — Clear IDs

> **Parent US**: [US-004-10](US-004-10-clear-ids.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 4S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-002-01 (AutoCAD connected — `IAutoCADService.IsConnected` must return true for COM operations)

## Acceptance Criteria
- [ ] AC-01: Clicking "Clear All IDs" clears column 8 across all layouts and tables in the active drawing
- [ ] AC-02: The clear operation iterates all non-Model layouts and all legal tables within each layout
- [ ] AC-03: For each table, column 8 is cleared for all rows except row 1 (the label row)
- [ ] AC-04: The operation is executed via `IGUIProxy.ExecuteInGuiAsync()` for STA thread safety
- [ ] AC-05: After clearing, the DataGrid DetailId column is updated to reflect empty values
- [ ] AC-06: A confirmation dialog is shown before clearing: "This will clear all Odoo IDs from the drawing. Continue?"
- [ ] AC-07: A success message is shown after clearing completes

---

## TASK-004-10-01: Add "Clear All IDs" button to extraction panel in BOQ Manager page

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Clear All IDs" `Button` in the extraction panel button bar (alongside "Extract from AutoCAD")
- Bind `Command="{Binding ClearAllIdsCommand}"`
- The button should be enabled only when AutoCAD is connected: use `CanExecute` from the command
- Style with a warning appearance (e.g., orange/amber border or icon) to indicate it is a destructive operation
- Position after the Extract button and before any Refresh button

### How to verify
- [ ] "Clear All IDs" button renders in the extraction panel (AC-01)
- [ ] Button is disabled when AutoCAD is not connected

---

## TASK-004-10-02: Implement ClearAllIdsCommand that calls IGUIProxy for COM clear operation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-10-05 |
| Blocks | TASK-004-10-06 |

### What to do
- Define `ClearAllIdsCommand` as `IAsyncRelayCommand` using `AsyncRelayCommand(ExecuteClearAllIdsAsync, CanExecuteClearAllIds)`
- `CanExecuteClearAllIds()` returns `IsAutoCADConnected && !IsExtracting && !IsPushing`
- In `ExecuteClearAllIdsAsync()`:
  1. Show confirmation dialog (from TASK-004-10-05): if user cancels, return immediately
  2. Call `_guiProxy.ExecuteInGuiAsync("clear_all_tables_id")` to execute the clear operation on the STA thread
  3. Check the `GUIProxyResponse` for success/failure
  4. On success: show a success message via `PushStatusText = "All IDs cleared successfully."` or a `MessageBox`
  5. On failure: show the error message from the response
- Log the clear operation: `_logger.LogInformation("Clear All IDs initiated")`

### How to verify
- [ ] ClearAllIdsCommand calls IGUIProxy for COM clear operation (AC-04)
- [ ] Success message is shown after clearing (AC-07)

---

## TASK-004-10-03: Implement clear_all_tables_id equivalent in IAutoCADService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-10-04 |

### What to do
- Add a method `ClearAllTablesId()` to `AutoCADService` (may also need to add to `IAutoCADService` interface)
- Implementation mirrors Python `clear_all_tables_id()`:
  1. Get all layouts via `GetLayouts()`, filter out Model space
  2. For each non-Model layout, switch to it and get all tables
  3. For each table, check if it is a legal table via `BOQProcessor.IsLegalTable()`
  4. For each legal table, iterate all rows:
     - Skip row 1 (label row)
     - For row 0: clear cell (0, 8) — this removes the `header_id`
     - For rows 2+: clear cell (i, 8) — this removes the `detail_id`
  5. Use `SetTableValue(tableHandle, row, 8, "")` to clear each cell
- Handle per-table errors gracefully: log and continue to next table

### How to verify
- [ ] All non-Model layouts are iterated (AC-02)
- [ ] Column 8 is cleared for all rows except row 1 in each legal table (AC-03)
- [ ] Model layout is excluded (AC-02)

---

## TASK-004-10-04: Register clear handler in IGUIProxy for STA thread execution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-004-10-03 |
| Blocks | TASK-004-10-02 |

### What to do
- Register a `"clear_all_tables_id"` handler with `IGUIProxy.RegisterHandler()` during application startup
- The handler calls `IAutoCADService.ClearAllTablesId()` (or the equivalent method) on the GUI thread
- Handler takes no parameters (clears all layouts in the active document)
- Returns a success/failure result with any error messages
- Ensure the handler executes synchronously on the STA thread

### How to verify
- [ ] `ExecuteInGuiAsync("clear_all_tables_id")` executes the clear operation on the GUI thread (AC-04)
- [ ] COM thread safety is maintained (AC-04)

---

## TASK-004-10-05: Add confirmation dialog before executing clear operation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-10-02 |

### What to do
- At the beginning of `ExecuteClearAllIdsAsync()`, show a `MessageBox` confirmation:
  - `MessageBox.Show("This will clear all Odoo IDs from the drawing. Continue?", "Confirm Clear", MessageBoxButton.YesNo, MessageBoxImage.Warning)`
- If result is `MessageBoxResult.No`, return without executing the clear operation
- Alternatively, use a dialog service interface (`IDialogService`) for testability:
  - `bool confirmed = await _dialogService.ConfirmAsync("This will clear all Odoo IDs from the drawing. Continue?", "Confirm Clear")`
- This prevents accidental clearing of IDs from the drawing

### How to verify
- [ ] Confirmation dialog appears before clearing (AC-06)
- [ ] Clear does not proceed if user cancels (AC-06)

---

## TASK-004-10-06: Update DataGrid DetailId values to empty after clear completes

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-10-02 |
| Blocks | None |

### What to do
- After the clear operation succeeds, update the in-memory data to reflect the cleared state:
  - Iterate all `BOQEntryRow` in `AllEntries` and set `DetailId = string.Empty`
  - Iterate all `BOQLayoutGroup` in `LayoutGroups` and set `HeaderId = string.Empty`
- Call `OnPropertyChanged` on each modified entry to refresh the DataGrid display
- This ensures the DataGrid reflects the actual state of the AutoCAD drawing after clearing
- If no data has been extracted yet (AllEntries is empty), skip the update but still show the success message

### How to verify
- [ ] After clearing, DataGrid DetailId column shows empty values (AC-05)
- [ ] HeaderId on layout groups is cleared (AC-05)

---

## Dependency Graph
```
TASK-004-10-03 (ClearAllTablesId Service)
       │
       └──▶ TASK-004-10-04 (IGUIProxy Clear Handler)

TASK-004-10-05 (Confirmation Dialog) ──────┐
TASK-004-10-04 (IGUIProxy Clear Handler) ──┤
                                            ▼
                                TASK-004-10-02 (ClearAllIdsCommand)
                                            │
                                            └──▶ TASK-004-10-06 (Update DataGrid)

TASK-004-10-01 (Clear Button XAML) ─── independent

Tasks 01, 03, 05 can begin in parallel.
Task 04 depends on 03. Task 02 depends on 04 and 05.
Task 06 depends on 02.
```
