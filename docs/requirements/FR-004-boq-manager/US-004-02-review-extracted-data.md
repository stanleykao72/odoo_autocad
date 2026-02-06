# US-004-02: Review Extracted Data

## User Story
**As a** CAD Engineer,
**I want to** see extracted BOQ data in a structured grid before pushing,
**So that** I can review and verify the data before it goes to Odoo.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Extracted BOQ data is displayed in a WPF DataGrid with columns: Layout Name, Position, Product No, Width, Height, Length, Thickness, Qty, Description, Detail ID, and Validation Status indicator
- [ ] AC-02: DataGrid rows are grouped by layout name, with each group showing a collapsible header containing layout metadata (PR No, Project Name, Job Working Plan Name, Product Name, Spec, Surface Treatment, Color)
- [ ] AC-03: A summary panel shows: Total Layouts processed, Total Items extracted, Valid Items count, Invalid Items count, and Skipped Items count
- [ ] AC-04: Rows with validation errors have a red background, warnings have a yellow background, and valid rows have a default/green background
- [ ] AC-05: Summary counts update automatically after extraction and validation via property change notifications
- [ ] AC-06: The DataGrid is read-only and displays data exactly as extracted (after MText unformatting)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-008 | Display extracted BOQ data in a DataGrid with specified columns and Validation Status indicator | Must |
| FR-004-009 | Group rows by layout name with collapsible header showing layout metadata | Should |
| FR-004-010 | Display summary panel with Total Layouts, Total Items, Valid, Invalid, and Skipped counts | Must |
| FR-004-011 | Visually distinguish rows by validation status with colored backgrounds | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-02-01 | Define DataGrid XAML with all columns and CollectionViewSource grouping by LayoutName | `Views/Pages/BOQManagerPage.xaml` | M |
| TASK-004-02-02 | Implement GroupStyle with collapsible Expander showing layout metadata in group header | `Views/Pages/BOQManagerPage.xaml` | M |
| TASK-004-02-03 | Implement summary panel using UniformGrid with 5 cells and colored backgrounds | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-02-04 | Create ValidationStatusToIconConverter for Status column DataTemplate (green check, red exclamation, yellow triangle) | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-02-05 | Bind AllEntries ObservableCollection to DataGrid ItemsSource with grouping | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-02-06 | Implement RowValidationStatusToBackgroundConverter for row-level coloring | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-02-07 | Wire summary count properties (TotalItems, ValidItems, WarningItems, ErrorItems, SkippedItems) with OnPropertyChanged notifications | `ViewModels/BOQManagerViewModel.cs` | S |

## Dependencies
- Depends on: US-004-01 (extraction must complete to populate grid)
- Blocks: US-004-03 (validation requires data in grid)

## Notes
- The DataGrid uses `CollectionViewSource` with `PropertyGroupDescription` on `LayoutName` to achieve the layout grouping.
- The Status column uses a `DataTemplate` with `{Binding ValidationStatus, Converter={StaticResource ValidationStatusToIconConverter}}` to show appropriate icons.
- Summary panel uses a `UniformGrid` with 5 cells. Each cell has a bold count and colored background: green for valid, yellow for warnings, red for errors, gray for skipped.
- The `BOQEntryRow` and `BOQLayoutGroup` models from the ViewModel section of FR-004 define the data structure.
