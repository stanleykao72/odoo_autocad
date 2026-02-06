# TASKS: US-004-06 — View Skipped Rows

> **Parent US**: [US-004-06](US-004-06-view-skipped-rows.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P2
> **Tasks**: 5 | **Effort**: 4S + 1M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-01 (Extraction process generates skip information — row-level and table-level skip counts)

## Acceptance Criteria
- [ ] AC-01: Rows skipped because both `qty` and `product_no` are empty/whitespace are counted and reported in the summary panel under "Skipped Items"
- [ ] AC-02: Tables skipped because they fail the legal table check (not 9 columns or missing HEADER_ID marker) are counted separately
- [ ] AC-03: The summary panel displays the total Skipped Items count with a gray background
- [ ] AC-04: A tooltip or expandable section on the Skipped count shows a breakdown: number of empty rows skipped, number of illegal tables skipped
- [ ] AC-05: The user can identify which layouts contained skipped rows or illegal tables

---

## TASK-004-06-01: Track skipped row count during extraction (empty qty + product_no rows)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-06-05 |

### What to do
- Add a `SkippedRowCount` property or counter to the `BOQGenerationResult` class
- During extraction in `ExtractItemsFromTable()` / `GetTableData()`, increment the counter each time `ShouldSkipRow()` returns true
- Include the skip reason in a `List<SkippedItemInfo>` collection on the result:
  - `LayoutName`: which layout the row was in
  - `RowIndex`: the table row index
  - `Reason`: "Empty qty and product_no"
- Return the total skipped row count in `BOQGenerationResult.SkippedItems`

### How to verify
- [ ] Empty rows (both qty and product_no empty) are counted and reported (AC-01)
- [ ] Skip count is available in the generation result

---

## TASK-004-06-02: Track skipped table count during extraction (illegal tables)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-06-05 |

### What to do
- Add a `SkippedTableCount` field or property to `BOQGenerationResult`
- During extraction, when `IsLegalTable(table)` returns false, increment the skipped table counter
- Add entries to the `SkippedItemInfo` list with:
  - `LayoutName`: which layout the table was in
  - `Reason`: "Table has {N} columns (expected 9)" or "Missing HEADER_ID marker at cell (0,7)"
- Track both categories separately: tables with wrong column count vs. tables missing HEADER_ID marker
- Note: Layouts named "Model" are excluded entirely and are NOT counted as skipped (per VR-004-004)

### How to verify
- [ ] Illegal tables (wrong column count or missing HEADER_ID) are counted separately (AC-02)
- [ ] Model layout is excluded, not counted as skipped (AC-02)

---

## TASK-004-06-03: Display SkippedItems in summary panel with gray background

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-06-04 |

### What to do
- Ensure the Skipped Items cell in the summary `UniformGrid` (from US-004-02 TASK-004-02-03) displays correctly:
  - Gray background (`#F3F4F6`)
  - Bold count bound to `{Binding SkippedItems}`
  - Sub-label "Skipped"
- The `SkippedItems` property in `BOQManagerViewModel` should be the total of skipped rows + skipped tables
- Ensure the property is updated after extraction completes via `UpdateSummaryCounts()`

### How to verify
- [ ] Summary panel shows Skipped Items with gray background (AC-03)
- [ ] Count reflects total skipped rows and skipped tables (AC-03)

---

## TASK-004-06-04: Add tooltip on Skipped count showing breakdown (empty rows vs illegal tables)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | TASK-004-06-03 |
| Blocks | None |

### What to do
- Add a `ToolTip` to the Skipped Items `Border` in the summary panel
- Bind the tooltip content to a `SkippedItemsTooltip` string property on the ViewModel
- The property should return a formatted string like: `"Empty rows: 5\nIllegal tables: 3"`
- Alternatively, use a `ToolTip` with a `StackPanel` containing two `TextBlock` elements bound to `SkippedRowCount` and `SkippedTableCount` respectively
- Add `SkippedRowCount` and `SkippedTableCount` observable properties to `BOQManagerViewModel` for the breakdown values
- When `SkippedItems` is 0, the tooltip can show "No rows or tables were skipped"

### How to verify
- [ ] Tooltip shows breakdown of empty rows vs illegal tables (AC-04)
- [ ] Both categories are separately displayed (AC-04)

---

## TASK-004-06-05: Store per-layout skip information for user reference

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-06-01, TASK-004-06-02 |
| Blocks | None |

### What to do
- Add a `SkippedItemsDetail` `ObservableCollection<SkippedItemInfo>` property to `BOQManagerViewModel`
- Define `SkippedItemInfo` class with: `LayoutName`, `ItemType` ("Row" or "Table"), `RowIndex` (nullable), `Reason`
- Populate from the `BOQGenerationResult` skip data after extraction completes
- Optionally display in a collapsible/expandable section below the summary panel or as a popup when clicking the Skipped count
- Group the skipped items by layout name so the user can identify which layouts had issues
- This enables users to navigate back to the drawing and fix the problematic layouts/tables

### How to verify
- [ ] Users can identify which layouts contained skipped rows or illegal tables (AC-05)
- [ ] Skipped items are grouped by layout name (AC-05)

---

## Dependency Graph
```
TASK-004-06-01 (Track Skipped Rows) ─────┐
                                          │
TASK-004-06-02 (Track Skipped Tables) ────┤
                                          ▼
                               TASK-004-06-05 (Per-Layout Skip Info)

TASK-004-06-03 (Summary Panel Display)
       │
       └──▶ TASK-004-06-04 (Tooltip Breakdown)

Tasks 01, 02, 03 can begin in parallel.
Task 05 depends on 01 and 02 (skip data needed).
Task 04 depends on 03 (panel must exist for tooltip).
```
