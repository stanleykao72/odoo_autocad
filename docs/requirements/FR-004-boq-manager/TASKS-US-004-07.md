# TASKS: US-004-07 — View Summary

> **Parent US**: [US-004-07](US-004-07-view-summary.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 6S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-01 (Extraction provides total and skipped counts)
- [ ] US-004-03 (Validation provides valid/warning/error counts)

## Acceptance Criteria
- [ ] AC-01: The summary panel displays Total Layouts processed
- [ ] AC-02: The summary panel displays Total Items extracted
- [ ] AC-03: The summary panel displays Valid Items count with green background
- [ ] AC-04: The summary panel displays Invalid Items count (Warnings with yellow background, Errors with red background)
- [ ] AC-05: The summary panel displays Skipped Items count with gray background
- [ ] AC-06: Summary counts are updated automatically after both extraction and validation complete
- [ ] AC-07: When no data has been extracted, the summary panel shows all zeros with an appropriate message

---

## TASK-004-07-01: Implement summary bar XAML using UniformGrid with 5 cells and color-coded backgrounds

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-07-02 |

### What to do
- Implement the summary bar as a `UniformGrid Columns="5"` inside a `Border` with margin and padding
- Cell 1: "Total" — neutral background `#E5E7EB`, bold count, sub-label "Items"
- Cell 2: "Valid" — green background `#D1FAE5`, bold count, sub-label "Valid"
- Cell 3: "Warnings" — yellow background `#FEF3C7`, bold count, sub-label "Warnings"
- Cell 4: "Errors" — red background `#FEE2E2`, bold count, sub-label "Errors"
- Cell 5: "Skipped" — gray background `#F3F4F6`, bold count, sub-label "Skipped"
- Add a "Layouts: {count}" label above the bar bound to `{Binding LayoutsFound}`
- When all counts are 0, show a subtle "No data extracted" message via a `DataTrigger` or `Visibility` binding

### How to verify
- [ ] Summary bar renders with 5 color-coded cells (AC-01 through AC-05)
- [ ] "No data extracted" message shown when all counts are zero (AC-07)

---

## TASK-004-07-02: Bind summary count properties to UI elements

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | TASK-004-07-01 |
| Blocks | None |

### What to do
- Bind each cell's count `TextBlock` to the corresponding ViewModel property:
  - Cell 1: `{Binding TotalItems}`
  - Cell 2: `{Binding ValidItems}`
  - Cell 3: `{Binding WarningItems}`
  - Cell 4: `{Binding ErrorItems}`
  - Cell 5: `{Binding SkippedItems}`
- Bind the layouts label to `{Binding LayoutsFound}`
- Ensure bindings use `Mode=OneWay` (default for TextBlock.Text) to update when properties change
- Set `FallbackValue="0"` on each binding for graceful handling before data is loaded

### How to verify
- [ ] Total Layouts count displays (AC-01)
- [ ] Total Items count displays (AC-02)
- [ ] Valid, Warning, Error, Skipped counts display with correct values (AC-03, AC-04, AC-05)

---

## TASK-004-07-03: Implement summary calculation method that tallies counts from AllEntries collection

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-07-04, TASK-004-07-05 |

### What to do
- Create `private void UpdateSummaryCounts()` method in `BOQManagerViewModel`
- Calculate: `TotalItems = AllEntries.Count`
- Calculate: `ValidItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Valid)`
- Calculate: `WarningItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Warning)`
- Calculate: `ErrorItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Error)`
- `SkippedItems` is set from the extraction result (not derived from `AllEntries` since skipped rows are not in the collection)
- Before validation runs, entries have `RowValidationStatus.Pending` — treat Pending as neither valid nor invalid for summary purposes
- Also update `HasValidationErrors` and `CanPush` based on the new counts

### How to verify
- [ ] Summary counts accurately reflect the state of AllEntries (AC-02, AC-03, AC-04, AC-05)
- [ ] Pending entries do not count as valid, warning, or error

---

## TASK-004-07-04: Trigger summary recalculation after extraction completes

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-07-03 |
| Blocks | None |

### What to do
- At the end of `ExecuteExtractAsync()`, after populating `AllEntries` and `LayoutGroups`, call `UpdateSummaryCounts()`
- Set `LayoutsFound` from the count of non-Model layouts discovered
- Set `SkippedItems` from the extraction result's `SkippedItems` property
- After extraction (before validation), entries will have `RowValidationStatus.Pending`, so `ValidItems`, `WarningItems`, `ErrorItems` will all be 0
- `TotalItems` and `SkippedItems` will reflect the extraction results

### How to verify
- [ ] Summary counts update after extraction completes (AC-06)
- [ ] TotalItems and SkippedItems are populated; validation counts are 0 pending validation (AC-06)

---

## TASK-004-07-05: Trigger summary recalculation after validation completes

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-07-03 |
| Blocks | None |

### What to do
- At the end of `ExecuteValidateAsync()`, after updating entry validation statuses, call `UpdateSummaryCounts()`
- After validation, entries will have `RowValidationStatus.Valid`, `Warning`, or `Error`, so all counts will be populated
- `TotalItems` remains unchanged from extraction; `ValidItems + WarningItems + ErrorItems` should equal `TotalItems`
- Update `PushStatusText` to reflect the new validation state (e.g., "Ready (42 valid, 3 warnings, 2 errors)")

### How to verify
- [ ] Summary counts update after validation completes with correct valid/warning/error breakdown (AC-06)
- [ ] ValidItems + WarningItems + ErrorItems == TotalItems after validation (AC-06)

---

## TASK-004-07-06: Add LayoutsFound count property and bind to summary panel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Declare `[ObservableProperty] private int _layoutsFound;` in `BOQManagerViewModel`
- Set `LayoutsFound` in `ExecuteExtractAsync()` from the count of non-Model layouts returned by the extraction
- This count represents the total number of layouts scanned (not the number with legal tables)
- Bind in the summary panel header: `"Layouts Found: {LayoutsFound}"`
- Default to 0 before extraction

### How to verify
- [ ] LayoutsFound shows the number of non-Model layouts (AC-01)
- [ ] Value is 0 before extraction and updates after extraction (AC-07)

---

## Dependency Graph
```
TASK-004-07-01 (Summary Bar XAML)
       │
       └──▶ TASK-004-07-02 (Bind Properties to UI)

TASK-004-07-03 (Calculation Method)
       │
       ├──▶ TASK-004-07-04 (Post-Extraction Trigger)
       └──▶ TASK-004-07-05 (Post-Validation Trigger)

TASK-004-07-06 (LayoutsFound Property) ─── independent

All tasks are S-sized. Tasks 01, 03, 06 can begin in parallel.
Task 02 depends on 01. Tasks 04, 05 depend on 03.
```
