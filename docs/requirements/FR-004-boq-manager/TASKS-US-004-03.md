# TASKS: US-004-03 — Validate Against Products

> **Parent US**: [US-004-03](US-004-03-validate-against-products.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 4S + 2M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-01 (Extraction must provide data to validate — `AllEntries` must be populated)
- [ ] US-003-01 (Odoo connected — `IOdooService.IsConnected` for product catalog queries)

## Acceptance Criteria
- [ ] AC-01: Clicking "Validate" checks all extracted BOQ entries against Odoo product catalog and business rules
- [ ] AC-02: Each entry's `product_no` is verified to map to a valid Odoo product ID (case-insensitive matching via local mapping cache first, then Odoo query)
- [ ] AC-03: Each entry's `qty` is verified to be a positive numeric value; non-numeric or negative values produce Error severity; zero produces Warning severity
- [ ] AC-04: A valid project context (non-zero `ProjectId` from `pr_no` lookup) is verified before push is allowed
- [ ] AC-05: Per-entry validation errors include entry index, field name, error message, and severity level (Error or Warning)
- [ ] AC-06: Rows with non-empty `product_no` but empty `qty` produce a Warning-severity validation error
- [ ] AC-07: Unresolved `product_no` values produce Error severity (or Warning if `AutoCreateProducts` is enabled)
- [ ] AC-08: Validation results update the row-level ValidationStatus and ValidationMessage properties on each BOQEntryRow
- [ ] AC-09: Summary counts (ValidItems, WarningItems, ErrorItems) are recalculated after validation

---

## TASK-004-03-01: Add "Validate All" button to BOQ Manager page bound to ValidateCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Validate All" `Button` in the push controls section of `BOQManagerPage.xaml`
- Bind `Command="{Binding ValidateCommand}"`
- Set `IsEnabled` to reflect `CanExecute`: validation requires `AllEntries.Count > 0` and `!IsValidating`
- Position the button before the "Push to Odoo" button in the horizontal button bar
- Add a `TextBlock` beside the button that shows "Validating..." when `IsValidating` is true, using `BoolToVisibilityConverter`

### How to verify
- [ ] "Validate All" button renders and is bound to ValidateCommand (AC-01)
- [ ] Button is disabled when no data is extracted or validation is in progress

---

## TASK-004-03-02: Implement ValidateCommand that calls IBOQProcessor.ValidateBOQAsync(entries)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-03-01 |
| Blocks | TASK-004-03-07, TASK-004-03-08 |

### What to do
- Define `ValidateCommand` as `IAsyncRelayCommand` using `AsyncRelayCommand(ExecuteValidateAsync, CanExecuteValidate)`
- `CanExecuteValidate()` returns `!IsValidating && AllEntries.Count > 0`
- In `ExecuteValidateAsync()`: set `IsValidating = true`, convert `AllEntries` to `IEnumerable<BOQEntry>` by mapping `BOQEntryRow` properties to `BOQEntry` fields
- Call `_boqProcessor.ValidateBOQAsync(entries)` to get `BOQValidationResult`
- Convert `BOQValidationResult.Errors` list to `ObservableCollection<BOQValidationErrorDisplay>` and set `ValidationErrors` property
- Set `HasValidationErrors = result.Errors.Any(e => e.Severity == "Error")`
- Update `CanPush` based on: `!HasValidationErrors && TotalItems > 0`
- Set `IsValidating = false` in a `finally` block
- Call `UpdateEntryValidationStatuses(result)` and `UpdateSummaryCounts()` at the end

### How to verify
- [ ] ValidateCommand calls `IBOQProcessor.ValidateBOQAsync()` with extracted entries (AC-01)
- [ ] Per-entry errors with index, field, message, severity are produced (AC-05)
- [ ] `HasValidationErrors` and `CanPush` are updated after validation

---

## TASK-004-03-03: Implement ValidateBOQAsync with product_no resolution (local cache then Odoo query)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-004-03-02 |

### What to do
- Enhance `ValidateBOQAsync()` in `BOQProcessor` to implement the full validation pipeline per FR-004 Section 8:
- For each entry, resolve `product_no` using two-tier lookup:
  1. Check `_productMappings` dictionary (case-insensitive via `StringComparer.OrdinalIgnoreCase`)
  2. If not found, call `_odooService.SearchProductsAsync(entry.ProductName)` and cache the result
- If product cannot be resolved and `BOQGenerationOptions.ValidateProducts` is true, add `BOQValidationError` with severity "Error" on field "ProductId" (VR-004-010)
- If `BOQGenerationOptions.AutoCreateProducts` is true, unresolved products produce severity "Warning" instead (VR-004-011)
- Validate project context: if `entry.ProjectId <= 0`, add error on field "ProjectId" (VR-004-012)
- Return `BOQValidationResult` with `IsValid = !errors.Any(e => e.Severity == "Error")`

### How to verify
- [ ] Product resolution checks local cache first, then queries Odoo (AC-02)
- [ ] Case-insensitive matching via OrdinalIgnoreCase (AC-02)
- [ ] Unresolved products produce Error or Warning based on options (AC-07)
- [ ] Project context validation blocks push when ProjectId is invalid (AC-04)

---

## TASK-004-03-04: Implement qty validation rules: positive numeric required, zero=Warning, negative/non-numeric=Error

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-03-03 |

### What to do
- Add qty validation logic to `ValidateBOQAsync()` applying VR-004-006 through VR-004-008:
- Parse the `qty` string using `decimal.TryParse()`:
  - If parse fails (non-numeric), add `BOQValidationError` with severity "Error", field "Quantity", message "Quantity must be a valid number" (VR-004-007)
  - If value is negative, add error with severity "Error", field "Quantity", message "Quantity must be greater than zero" (VR-004-007)
  - If value is zero, add error with severity "Warning", field "Quantity", message "Quantity is zero" (VR-004-008)
- If `product_no` is non-empty but `qty` is empty/whitespace, add warning with severity "Warning", field "Quantity", message "Product has no quantity specified" (VR-004-006)
- Include the `EntryIndex` in each error for row-level navigation

### How to verify
- [ ] Non-numeric qty produces Error severity (AC-03)
- [ ] Negative qty produces Error severity (AC-03)
- [ ] Zero qty produces Warning severity (AC-03)
- [ ] Non-empty product_no with empty qty produces Warning (AC-06)

---

## TASK-004-03-05: Implement project context validation via IOdooService.GetProjectAsync()

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-03-03 |

### What to do
- Add project context validation at the beginning of `ValidateBOQAsync()`:
- Check if a valid `ProjectId` (> 0) is available from the entry data
- If `pr_no` is available but `ProjectId` is not resolved, call `_odooService.SearchProjectsAsync(prNo)` to look up the project (equivalent to Python `get_project(pr_no)`)
- If no matching project is found, add a `BOQValidationError` with severity "Error", field "ProjectId", message "No valid project found for PR number '{pr_no}'" (VR-004-012, VR-004-013)
- This is a global validation (not per-entry) that blocks the entire push operation

### How to verify
- [ ] Valid ProjectId (> 0) is verified before push is allowed (AC-04)
- [ ] Missing or invalid project produces Error on field "ProjectId" (AC-04)

---

## TASK-004-03-06: Build BOQValidationResult/BOQValidationError model population with per-entry details

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-03-02 |

### What to do
- Ensure each `BOQValidationError` created during validation includes all required fields:
  - `EntryIndex`: the 0-based index of the entry in the collection (nullable for global errors)
  - `Field`: the property name being validated ("ProductId", "Quantity", "ProjectId")
  - `Message`: a user-readable error description
  - `Severity`: "Error" or "Warning"
- Aggregate all errors into `BOQValidationResult.Errors` list
- Set `BOQValidationResult.IsValid` to `true` only if no Error-severity items exist (Warnings alone do not block push)
- Log each validation error via `_logger?.LogWarning()` for structured logging

### How to verify
- [ ] Every validation error has EntryIndex, Field, Message, and Severity populated (AC-05)
- [ ] IsValid is false only when Error-severity issues exist (AC-05)

---

## TASK-004-03-07: Update BOQEntryRow.ValidationStatus and ValidationMessage after validation completes

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-03-02 |
| Blocks | None |

### What to do
- Create a private method `UpdateEntryValidationStatuses(BOQValidationResult result)` in `BOQManagerViewModel`
- First, set all entries to `RowValidationStatus.Valid` and clear `ValidationMessage`
- Then iterate `result.Errors` and for each error with a non-null `EntryIndex`:
  - Find the corresponding `BOQEntryRow` in `AllEntries` by index
  - Set `ValidationStatus` to `RowValidationStatus.Error` if severity is "Error", or `RowValidationStatus.Warning` if severity is "Warning"
  - Set `ValidationMessage` to the error message (concatenate if multiple errors on same row)
- If an entry has both Error and Warning, Error takes precedence
- Call `OnPropertyChanged` on each modified entry to trigger DataGrid row re-rendering

### How to verify
- [ ] Each BOQEntryRow's ValidationStatus reflects the validation result (AC-08)
- [ ] ValidationMessage contains the specific error description (AC-08)

---

## TASK-004-03-08: Recalculate summary counts (ValidItems, WarningItems, ErrorItems) post-validation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-03-02 |
| Blocks | None |

### What to do
- Call `UpdateSummaryCounts()` at the end of `ExecuteValidateAsync()` after `UpdateEntryValidationStatuses()` completes
- The method iterates `AllEntries` and counts entries by their `ValidationStatus`:
  - `ValidItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Valid)`
  - `WarningItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Warning)`
  - `ErrorItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Error)`
- Update `PushStatusText` to reflect readiness: e.g., "Ready (42 valid entries, 2 errors to resolve)" or "All entries valid - ready to push"
- Notify UI via `OnPropertyChanged` for all summary properties

### How to verify
- [ ] ValidItems, WarningItems, ErrorItems counts match the validation results (AC-09)
- [ ] PushStatusText reflects the current validation state (AC-09)

---

## Dependency Graph
```
TASK-004-03-04 (Qty Validation Rules)
       │
       └──▶ TASK-004-03-03 (Product Resolution Validation)
TASK-004-03-05 (Project Context Validation)
       │
       └──▶ TASK-004-03-03 (Product Resolution Validation)
TASK-004-03-06 (Error Model Population)
       │
       └──▶ TASK-004-03-02 (ValidateCommand ViewModel)

TASK-004-03-01 (Validate Button XAML)
       │
       └──▶ TASK-004-03-02 (ValidateCommand ViewModel)
                   │
                   ├──▶ TASK-004-03-07 (Update Entry Statuses)
                   └──▶ TASK-004-03-08 (Recalculate Summary)

Tasks 01, 04, 05, 06 can begin in parallel.
Tasks 04 and 05 feed into 03 (core validation logic).
Tasks 07 and 08 depend on 02 (ViewModel command).
```
