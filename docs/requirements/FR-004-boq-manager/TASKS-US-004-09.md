# TASKS: US-004-09 — View Validation Errors

> **Parent US**: [US-004-09](US-004-09-view-validation-errors.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 4S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-03 (Validation must produce errors to display — `BOQValidationResult.Errors` must be populated)

## Acceptance Criteria
- [ ] AC-01: Validation results are displayed in a dedicated panel below the DataGrid
- [ ] AC-02: Each validation error shows: severity icon, row reference (e.g., "Row 3 (Sheet-1)"), field name, and error message
- [ ] AC-03: Clicking a validation error selects and highlights the corresponding row in the DataGrid
- [ ] AC-04: Each validation error entry provides inline action buttons: "Map Product" (for product errors), "Edit" (for data errors), "Ignore" (to dismiss warnings)
- [ ] AC-05: Error-severity entries are displayed with a red severity icon; Warning-severity entries with a yellow triangle icon
- [ ] AC-06: The validation results panel shows a count header (e.g., "2 errors, 3 warnings")
- [ ] AC-07: The panel is scrollable when there are many validation errors

---

## TASK-004-09-01: Create validation results panel XAML as a ListView of BOQValidationErrorDisplay items below the DataGrid

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-09-02, TASK-004-09-07 |

### What to do
- Add a validation results section below the DataGrid in `BOQManagerPage.xaml`
- Use a `Border` container with a header area and a `ListView` content area
- Bind `ListView.ItemsSource` to `{Binding ValidationErrors}` (`ObservableCollection<BOQValidationErrorDisplay>`)
- Set `ListView.MaxHeight` to a reasonable value (e.g., 200) and enable `ScrollViewer.VerticalScrollBarVisibility="Auto"` for scrollability
- Add `Visibility` binding: visible when `ValidationErrors.Count > 0` using a `CountToVisibilityConverter` or `DataTrigger`
- Set `ListView.SelectedItem` bound to a `SelectedValidationError` property for click handling

### How to verify
- [ ] Validation results panel renders below the DataGrid (AC-01)
- [ ] Panel is scrollable when many errors exist (AC-07)
- [ ] Panel is hidden when no validation errors exist

---

## TASK-004-09-02: Implement DataTemplate for validation error items with severity icon, row reference, field, message, and action buttons

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | M |
| Depends On | TASK-004-09-01 |
| Blocks | None |

### What to do
- Define a `DataTemplate` for `BOQValidationErrorDisplay` in the ListView
- Layout as a horizontal `Grid` with columns:
  1. **Severity Icon**: `TextBlock` with icon bound to `{Binding Severity, Converter={StaticResource SeverityToIconConverter}}` — red `!` for Error, yellow triangle for Warning
  2. **Row Reference**: `TextBlock` bound to `{Binding RowReference}` (e.g., "Row 3 (Sheet-1)")
  3. **Field**: `TextBlock` bound to `{Binding Field}` in a subtle style
  4. **Message**: `TextBlock` bound to `{Binding Message}` with `TextWrapping="Wrap"`
  5. **Action Buttons**: `StackPanel Orientation="Horizontal"` containing context-sensitive buttons:
     - "Map Product" button visible when `Field == "ProductId"`, bound to `OpenProductMappingCommand` with `CommandParameter`
     - "Ignore" button visible when `Severity == "Warning"`, bound to `IgnoreWarningCommand`
- Create `SeverityToIconConverter` (`IValueConverter`) that maps "Error" -> red icon, "Warning" -> yellow icon

### How to verify
- [ ] Each error shows severity icon, row reference, field, and message (AC-02)
- [ ] Error severity shows red icon, Warning shows yellow triangle (AC-05)
- [ ] Inline action buttons are context-sensitive (AC-04)

---

## TASK-004-09-03: Implement NavigateToErrorCommand that selects/scrolls to the corresponding DataGrid row

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Define `NavigateToErrorCommand` as `IRelayCommand<BOQValidationErrorDisplay>` in `BOQManagerViewModel`
- When executed, use `error.EntryIndex` to find the corresponding `BOQEntryRow` in `AllEntries`
- Set a `SelectedEntryIndex` or `SelectedEntry` observable property to the target entry
- The DataGrid `SelectedItem` is bound to `{Binding SelectedEntry, Mode=TwoWay}` to highlight the row
- For scrolling: this may require code-behind in `BOQManagerPage.xaml.cs`:
  - Subscribe to a `ScrollToEntryRequested` event from the ViewModel
  - In the handler, call `DataGrid.ScrollIntoView(selectedItem)` and `DataGrid.SelectedItem = selectedItem`
- Also handle clicking directly on a ListView item (via `SelectionChanged` or `MouseDoubleClick`) to trigger navigation
- Wire the ListView item click to invoke `NavigateToErrorCommand` with the selected error

### How to verify
- [ ] Clicking an error in the validation panel selects the corresponding DataGrid row (AC-03)
- [ ] DataGrid scrolls to make the selected row visible (AC-03)

---

## TASK-004-09-04: Populate ValidationErrors ObservableCollection from IBOQProcessor.ValidateBOQAsync() results

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `ExecuteValidateAsync()`, after receiving `BOQValidationResult`, convert each `BOQValidationError` to `BOQValidationErrorDisplay`:
  - `EntryIndex = error.EntryIndex`
  - `LayoutName = AllEntries[error.EntryIndex.Value].LayoutName` (if index is valid)
  - `RowReference = $"Row {error.EntryIndex + 1} ({layoutName})"` (1-based for user display)
  - `Field = error.Field`
  - `Message = error.Message`
  - `Severity = error.Severity`
- Clear `ValidationErrors` collection and repopulate with the new display entries
- Handle null `EntryIndex` for global errors (e.g., project context): set `RowReference = "Global"` and `LayoutName = ""`

### How to verify
- [ ] ValidationErrors collection is populated from validation results (AC-01)
- [ ] Each error display has all required fields: index, layout, row ref, field, message, severity (AC-02)

---

## TASK-004-09-05: Implement "Map Product" inline action that opens ProductMappingDialog for the affected entry

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Wire the "Map Product" button in the validation error `DataTemplate` to `OpenProductMappingCommand`
- The `CommandParameter` passes the `BOQValidationErrorDisplay` item, from which the `EntryIndex` is used to find the corresponding `BOQEntryRow`
- Retrieve `entry.ProductNo` and pass it as the `AutoCADProductName` to the `ProductMappingDialog`
- After successful mapping, the error should be cleared from the `ValidationErrors` collection (handled by US-004-08 TASK-004-08-08)
- Only show the "Map Product" button when `error.Field == "ProductId"` using a `DataTrigger` or `Visibility` binding with a `FieldToVisibilityConverter`

### How to verify
- [ ] "Map Product" button opens the ProductMappingDialog with the correct product name (AC-04)
- [ ] Button is only visible for product-related errors (AC-04)

---

## TASK-004-09-06: Implement "Ignore" inline action that dismisses Warning-severity entries

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Define `IgnoreWarningCommand` as `IRelayCommand<BOQValidationErrorDisplay>` in `BOQManagerViewModel`
- When executed, remove the warning entry from `ValidationErrors` collection
- Update the corresponding `BOQEntryRow.ValidationStatus` to `RowValidationStatus.Valid` (since the warning is being dismissed)
- Recalculate summary counts via `UpdateSummaryCounts()` — `WarningItems` should decrease by 1, `ValidItems` increase by 1
- The "Ignore" button is only visible when `error.Severity == "Warning"` using a `DataTrigger`
- Log the dismissal: `_logger.LogInformation("Warning ignored for entry {Index}: {Message}", error.EntryIndex, error.Message)`

### How to verify
- [ ] "Ignore" button is only visible for Warning-severity entries (AC-04)
- [ ] Clicking "Ignore" removes the warning from the list and updates the entry status (AC-04)

---

## TASK-004-09-07: Add error/warning count header to the validation results panel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | TASK-004-09-01 |
| Blocks | None |

### What to do
- Add a header `TextBlock` above the ListView in the validation results panel
- Bind to a computed `ValidationSummaryText` property on the ViewModel (e.g., "2 errors, 3 warnings")
- Implement `ValidationSummaryText` as a computed property that counts errors and warnings in `ValidationErrors`:
  - `$"{errorCount} error{(errorCount != 1 ? "s" : "")}, {warningCount} warning{(warningCount != 1 ? "s" : "")}"`
- Update the property whenever `ValidationErrors` changes (subscribe to `CollectionChanged` or recalculate in `UpdateSummaryCounts()`)
- Style the header with bold text and appropriate margins

### How to verify
- [ ] Count header shows "X errors, Y warnings" (AC-06)
- [ ] Count updates when errors are resolved or warnings dismissed (AC-06)

---

## Dependency Graph
```
TASK-004-09-01 (Panel XAML Container)
       │
       ├──▶ TASK-004-09-02 (Error Item DataTemplate)
       └──▶ TASK-004-09-07 (Count Header)

TASK-004-09-03 (NavigateToError Command) ─── independent
TASK-004-09-04 (Populate ValidationErrors) ─── independent
TASK-004-09-05 (Map Product Action) ─── independent
TASK-004-09-06 (Ignore Warning Action) ─── independent

Task 01 is the foundation for the XAML container.
Tasks 02 and 07 depend on 01 (panel must exist).
Tasks 03, 04, 05, 06 can begin in parallel (ViewModel logic, no XAML dependency).
```
