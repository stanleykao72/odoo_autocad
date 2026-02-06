# US-004-09: View Validation Errors

## User Story
**As a** CAD Engineer,
**I want to** see validation errors with row-level detail,
**So that** I can navigate directly to the problematic entry.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Validation results are displayed in a dedicated panel below the DataGrid
- [ ] AC-02: Each validation error shows: severity icon, row reference (e.g., "Row 3 (Sheet-1)"), field name, and error message
- [ ] AC-03: Clicking a validation error selects and highlights the corresponding row in the DataGrid
- [ ] AC-04: Each validation error entry provides inline action buttons: "Map Product" (for product errors), "Edit" (for data errors), "Ignore" (to dismiss warnings)
- [ ] AC-05: Error-severity entries are displayed with a red severity icon; Warning-severity entries with a yellow triangle icon
- [ ] AC-06: The validation results panel shows a count header (e.g., "2 errors, 3 warnings")
- [ ] AC-07: The panel is scrollable when there are many validation errors

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-016 | Produce per-entry error details including entry index, field name, error message, and severity level | Must |
| FR-004-017 | Display validation results in panel below DataGrid with clickable errors that highlight corresponding row | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-09-01 | Create validation results panel XAML as a ListView of BOQValidationErrorDisplay items below the DataGrid | `Views/Pages/BOQManagerPage.xaml` | M |
| TASK-004-09-02 | Implement DataTemplate for validation error items with severity icon, row reference, field, message, and action buttons | `Views/Pages/BOQManagerPage.xaml` | M |
| TASK-004-09-03 | Implement NavigateToErrorCommand that selects/scrolls to the corresponding DataGrid row when an error is clicked | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-09-04 | Populate ValidationErrors ObservableCollection from IBOQProcessor.ValidateBOQAsync() results | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-09-05 | Implement "Map Product" inline action that opens ProductMappingDialog for the affected entry | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-09-06 | Implement "Ignore" inline action that dismisses Warning-severity entries | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-09-07 | Add error/warning count header to the validation results panel | `Views/Pages/BOQManagerPage.xaml` | S |

## Dependencies
- Depends on: US-004-03 (validation must produce errors to display)
- Blocks: None

## Notes
- The validation results panel is a `ListView` bound to `ValidationErrors` (`ObservableCollection<BOQValidationErrorDisplay>`).
- Each item uses a `DataTemplate` showing: severity icon (from a converter), the `RowReference` string (e.g., "Row 3 (Sheet-1)"), the `Field` name, and the `Message`.
- The `NavigateToErrorCommand` is a `RelayCommand<BOQValidationErrorDisplay>` that uses `EntryIndex` to find and select the corresponding row in the DataGrid. This may require code-behind in `BOQManagerPage.xaml.cs` to programmatically scroll the DataGrid.
- Action buttons are context-sensitive: "Map Product" appears only for product-related errors (field = "ProductId"), "Edit" for data errors, "Ignore" for warnings.
- The `BOQValidationErrorDisplay` model includes `EntryIndex`, `LayoutName`, `RowReference`, `Field`, `Message`, and `Severity` properties.
