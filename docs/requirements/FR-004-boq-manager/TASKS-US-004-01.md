# TASKS: US-004-01 — Extract BOQ Data

> **Parent US**: [US-004-01](US-004-01-extract-boq-data.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 4S + 2M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] US-002-01 (AutoCAD connected — `IAutoCADService.IsConnected` must return true)
- [ ] FR-008 (UI Framework — WPF/MVVM infrastructure and CommunityToolkit.Mvvm must be in place)

## Acceptance Criteria
- [ ] AC-01: Clicking "Extract from AutoCAD" iterates all non-Model layouts in the active document and collects attribute block values and table data
- [ ] AC-02: Only legal tables (exactly 9 columns with `HEADER_ID` at cell (0, 7)) are processed; others are silently skipped
- [ ] AC-03: Attribute block values are extracted for all required tags: `pr_no`, `project_name`, `job_working_plan_name`, `product_name`, `product_catelog`, `spec`, `surface_treatment`, `operation_flow`, `color_name`, `color_no`
- [ ] AC-04: Table rows where both `qty` (column 6) and `product_no` (column 1) are empty/whitespace after MText unformatting are skipped
- [ ] AC-05: MText formatting codes are stripped from all cell values using `MTextFormatter.UnFormat()` before populating the data grid
- [ ] AC-06: The `header_id` is read from cell (row 0, column 8) for each legal table
- [ ] AC-07: A progress indicator displays the current layout being processed (e.g., "Layout 3/5")
- [ ] AC-08: All AutoCAD COM operations are routed through `IGUIProxy.ExecuteInGuiAsync()` for STA thread safety
- [ ] AC-09: If AutoCAD is not connected, the Extract button is disabled and a message is shown: "AutoCAD is not connected. Please connect to AutoCAD before extracting BOQ data."
- [ ] AC-10: If no active document is found, the message "No active AutoCAD document found. Please open a drawing file." is shown

---

## TASK-004-01-01: Add "Extract from AutoCAD" button to BOQ Manager page with binding to ExtractCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-01-02 |

### What to do
- Create the `BOQManagerPage.xaml` WPF Page in namespace `OdooAutoCAD.App.Views.Pages`
- Add an Extraction Panel section at the top with `StackPanel` containing project context labels (CurrentPRNo, CurrentProjectName, CurrentDocumentName)
- Add a horizontal button bar with "Extract from AutoCAD" `Button` bound to `{Binding ExtractCommand}`
- Set `IsEnabled` binding to `{Binding IsAutoCADConnected}` so the button is disabled when AutoCAD is disconnected
- Add a `TextBlock` for connection warning message bound to `AutoCADConnectionMessage` with `Visibility` via `BoolToVisibilityConverter` on `!IsAutoCADConnected`
- Add corresponding `BOQManagerPage.xaml.cs` code-behind that sets `DataContext` to `BOQManagerViewModel` resolved from DI

### How to verify
- [ ] "Extract from AutoCAD" button renders and is bound to ExtractCommand (AC-01)
- [ ] Button is disabled when AutoCAD is not connected, showing the warning message (AC-09)

---

## TASK-004-01-02: Implement ExtractCommand in ViewModel that calls IGUIProxy for COM extraction

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | L |
| Depends On | TASK-004-01-01 |
| Blocks | TASK-004-01-06 |

### What to do
- Create `BOQManagerViewModel.cs` in namespace `OdooAutoCAD.App.ViewModels`, extending `ObservableObject`
- Inject `IBOQProcessor`, `IAutoCADService`, `IOdooService`, `IGUIProxy`, and `ILogger<BOQManagerViewModel>` via constructor
- Define `ExtractCommand` as `IAsyncRelayCommand` using `AsyncRelayCommand(ExecuteExtractAsync, CanExecuteExtract)`
- `CanExecuteExtract()` returns `!IsExtracting && IsAutoCADConnected`
- In `ExecuteExtractAsync()`: set `IsExtracting = true`, call `_guiProxy.ExecuteInGuiAsync("extract_layouts_values")` to execute COM on STA thread
- Parse the `GUIProxyResponse.Result` (expected `List<BOQLayoutGroup>`) and populate `LayoutGroups` and `AllEntries` `ObservableCollection` properties
- Build `BOQEntryRow` objects from the response, applying `MTextFormatter.UnFormat()` to all cell values
- Handle the "no active document" case by checking response for null/empty result and setting status message
- Set `IsExtracting = false` in a `finally` block
- Define all observable properties from FR-004 Section 6 ViewModel: `IsExtracting`, `ExtractionProgress`, `ExtractionProgressPercent`, `LayoutsFound`, `LayoutGroups`, `AllEntries`, and summary count properties

### How to verify
- [ ] ExtractCommand triggers `IGUIProxy.ExecuteInGuiAsync("extract_layouts_values")` for COM thread safety (AC-08)
- [ ] Extracted data populates `AllEntries` collection after successful extraction (AC-01)
- [ ] No active document case shows appropriate message (AC-10)

---

## TASK-004-01-03: Implement GetLayoutsValues() in IAutoCADService to iterate layouts, validate tables, and extract data

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | L |
| Depends On | TASK-004-01-07 |
| Blocks | TASK-004-01-05 |

### What to do
- Enhance the existing `GetLayoutsValues()` method in `AutoCADService` to implement the full Python `get_layouts_values()` equivalent
- Iterate all layouts from `GetLayouts()`, filtering out any where `IsModelSpace == true`
- For each non-Model layout, call `GetLayoutAttributeBlockValues(layout, tagList)` to extract attribute block values for the 10-tag list: `pr_no`, `project_name`, `job_working_plan_name`, `product_name`, `product_catelog`, `spec`, `surface_treatment`, `operation_flow`, `color_name`, `color_no`
- For each layout, call `GetLayoutTableBlocks(blocks)` to find `AcDbTable` objects, then filter with `CheckLegalTable(table)` for 9-column + HEADER_ID validation
- For legal tables, call `GetTableData(table)` to extract detail rows, applying `MTextFormatter.UnFormat()` to each cell value
- Read `header_id` from cell (row 0, column 8) for each legal table
- Return populated `LayoutData` with all extracted values structured as `BOQLayoutGroup` objects

### How to verify
- [ ] All non-Model layouts are iterated and attribute blocks extracted for 10 tags (AC-01, AC-03)
- [ ] Only legal tables (9 columns, HEADER_ID at (0,7)) are processed (AC-02)
- [ ] header_id is read from cell (0, 8) for each legal table (AC-06)

---

## TASK-004-01-04: Implement MTextFormatter.UnFormat() static method for stripping MText formatting codes

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Utilities/MTextFormatter.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-01-03, TASK-004-01-08 |

### What to do
- Create `MTextFormatter.cs` in namespace `OdooAutoCAD.Core.Utilities` as a static utility class
- Implement `public static string UnFormat(string mtext)` matching the Python `LM_UnFormat()` regex chain:
  1. Replace `\\\\` with ASCII 032 (space) placeholder
  2. Replace `\\P`, `\n`, `\t` with space
  3. Strip formatting commands: `\A`, `\C`, `\c`, `\F`, `\f`, `\H`, `\L`, `\l`, `\O`, `\o`, `\p`, `\Q`, `\T`, `\W` and their arguments using regex patterns like `\\[AaCcFfHhLlOoPpQqTtWw][^;]*;`
  4. Strip `\S` stacking expressions (e.g., `\S1/2;` -> remove entire match)
  5. Remove remaining escape characters `\\` and braces `{}`
  6. Trim the result
- Handle null/empty input by returning `string.Empty`
- Use `System.Text.RegularExpressions.Regex` with compiled patterns for performance

### How to verify
- [ ] MText formatting codes are stripped from cell values (AC-05)
- [ ] Unit tests verify each regex substitution step matches Python `LM_UnFormat()` output

---

## TASK-004-01-05: Register extraction handler in IGUIProxy for STA thread execution

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | M |
| Depends On | TASK-004-01-03 |
| Blocks | TASK-004-01-02 |

### What to do
- In `App.xaml.cs` or a dedicated startup registration class, register the `"extract_layouts_values"` handler with `IGUIProxy.RegisterHandler()`
- The handler should call `IAutoCADService.GetLayoutsValues()` on the GUI thread and return the result
- Also register the `"extract_autocad_parameters"` handler (used by `BOQProcessor.GenerateBOQForProjectAsync()`) if not already registered
- Ensure the handlers execute synchronously within the STA thread context (wrap in `Task.FromResult()` if needed)
- Register `BOQManagerViewModel` as singleton in the DI container: `services.AddSingleton<BOQManagerViewModel>()`
- Wire `IGUIProxy.Start()` call in the application startup sequence and `IGUIProxy.Stop()` in shutdown

### How to verify
- [ ] `ExecuteInGuiAsync("extract_layouts_values")` successfully invokes `IAutoCADService.GetLayoutsValues()` on the GUI thread (AC-08)
- [ ] COM thread conflicts are avoided during extraction (AC-08)

---

## TASK-004-01-06: Add progress reporting with IProgress<T> pattern during layout iteration

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-01-02 |
| Blocks | None |

### What to do
- Create a `Progress<(int current, int total, string message)>` instance in `ExecuteExtractAsync()`
- Pass the progress reporter to the extraction pipeline so it can report after each layout is processed
- In the progress callback, update `ExtractionProgress` (e.g., `"Layout 3/5 extracting..."`) and `ExtractionProgressPercent` (e.g., `3.0 / 5.0`)
- Update `LayoutsFound` with the total number of non-Model layouts discovered
- Ensure progress updates are marshalled to the UI thread via the `Progress<T>` pattern (which automatically captures `SynchronizationContext`)

### How to verify
- [ ] During extraction, progress text shows "Layout X/Y extracting..." (AC-07)
- [ ] ExtractionProgressPercent updates from 0.0 to 1.0 as layouts are processed (AC-07)

---

## TASK-004-01-07: Implement chk_legal_table equivalent: validate 9 columns and HEADER_ID marker

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-01-03 |

### What to do
- Add a static method `public static bool IsLegalTable(TableData table)` to `BOQProcessor` (or a new `BOQTableValidator` utility class in `OdooAutoCAD.Core.BOQ`)
- Check that `table.ColumnCount == 9` (matching Python `chk_legal_table` column count check)
- Check that `table.Cells.Count > 0` and `table.Cells[0].Count >= 8` and `table.Cells[0][7] == "HEADER_ID"` (matching Python cell (0,7) check)
- Return `true` only if both conditions pass; `false` otherwise
- Apply `MTextFormatter.UnFormat()` to the cell value at (0,7) before comparison in case the marker has formatting codes

### How to verify
- [ ] Tables with exactly 9 columns and "HEADER_ID" at cell (0,7) pass validation (AC-02)
- [ ] Tables with wrong column count or missing marker are silently skipped (AC-02)

---

## TASK-004-01-08: Implement row skip logic for empty qty + product_no rows

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | TASK-004-01-04 |
| Blocks | None |

### What to do
- Add a static method `public static bool ShouldSkipRow(List<string> rowCells)` to `BOQProcessor` (or the `BOQTableValidator` utility)
- Extract `product_no` from `rowCells[1]` and `qty` from `rowCells[6]` (0-indexed)
- Apply `MTextFormatter.UnFormat()` to both values
- Return `true` if both values are null, empty, or whitespace-only after unformatting
- Integrate this check into the table data extraction loop in `GetTableData()` or `ExtractItemsFromTable()`
- Increment a skip counter for reporting purposes

### How to verify
- [ ] Rows with empty qty AND empty product_no (after MText unformatting) are skipped (AC-04)
- [ ] Rows with at least one non-empty value (qty or product_no) are included in extraction (AC-04)

---

## Dependency Graph
```
TASK-004-01-04 (MTextFormatter)
       │
       ├──▶ TASK-004-01-07 (Legal Table Check)
       │         │
       │         └──▶ TASK-004-01-03 (GetLayoutsValues Implementation)
       │                   │
       │                   └──▶ TASK-004-01-05 (IGUIProxy Handler Registration)
       │                             │
       │                             └──▶ TASK-004-01-02 (ExtractCommand ViewModel)
       │                                       │
       │                                       └──▶ TASK-004-01-06 (Progress Reporting)
       │
       └──▶ TASK-004-01-08 (Row Skip Logic)

TASK-004-01-01 (Extract Button XAML)
       │
       └──▶ TASK-004-01-02 (ExtractCommand ViewModel)

Tasks 01-01 and 04 can begin in parallel (no interdependency).
Task 04 (MTextFormatter) is foundational and should be implemented first.
```
