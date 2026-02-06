# TASKS: US-004-02 — Review Extracted Data

> **Parent US**: [US-004-02](US-004-02-review-extracted-data.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P1
> **Tasks**: 7 | **Effort**: 5S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-01 (Extraction must complete to populate the grid with data)

## Acceptance Criteria
- [ ] AC-01: Extracted BOQ data is displayed in a WPF DataGrid with columns: Layout Name, Position, Product No, Width, Height, Length, Thickness, Qty, Description, Detail ID, and Validation Status indicator
- [ ] AC-02: DataGrid rows are grouped by layout name, with each group showing a collapsible header containing layout metadata (PR No, Project Name, Job Working Plan Name, Product Name, Spec, Surface Treatment, Color)
- [ ] AC-03: A summary panel shows: Total Layouts processed, Total Items extracted, Valid Items count, Invalid Items count, and Skipped Items count
- [ ] AC-04: Rows with validation errors have a red background, warnings have a yellow background, and valid rows have a default/green background
- [ ] AC-05: Summary counts update automatically after extraction and validation via property change notifications
- [ ] AC-06: The DataGrid is read-only and displays data exactly as extracted (after MText unformatting)

---

## TASK-004-02-01: Define DataGrid XAML with all columns and CollectionViewSource grouping by LayoutName

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-02-02, TASK-004-02-05 |

### What to do
- Add a `CollectionViewSource` in `BOQManagerPage.xaml` resources with `Source="{Binding AllEntries}"` and a `PropertyGroupDescription` on `LayoutName`
- Define a `DataGrid` with `IsReadOnly="True"` and `ItemsSource="{Binding Source={StaticResource GroupedEntries}}"`
- Add `DataGridTextColumn` bindings for all 11 columns: LayoutName, Position, ProductNo, Width, Height, Length, Thickness, Qty, Description, DetailId
- Add a `DataGridTemplateColumn` for the Validation Status indicator using a `DataTemplate` with a `TextBlock` or icon bound to `ValidationStatus`
- Set `AutoGenerateColumns="False"` and configure column widths (Position=50, ProductNo=100, dimensions=70 each, Qty=60, Description=*, DetailId=80, Status=40)
- Enable alternating row colors for readability

### How to verify
- [ ] DataGrid renders with all specified columns (AC-01)
- [ ] DataGrid is read-only and cannot be edited (AC-06)

---

## TASK-004-02-02: Implement GroupStyle with collapsible Expander showing layout metadata in group header

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | M |
| Depends On | TASK-004-02-01 |
| Blocks | None |

### What to do
- Add a `DataGrid.GroupStyle` to the DataGrid containing a `GroupStyle` with a `ContainerStyle` that uses an `Expander`
- The `Expander.Header` should display a `StackPanel` with: layout name (bold), PR No, Project Name, Job Working Plan Name, Product Name, Spec, Surface Treatment, and Color — bound from the `BOQLayoutGroup` properties
- Use `{Binding Name}` for the group key (LayoutName) and `{Binding ItemCount}` for the entry count within the group
- Retrieve layout metadata from the `BOQLayoutGroup` matching the group name — consider using a `MultiBinding` with a converter or binding via the `DataContext` of the group container
- Set `IsExpanded="True"` by default so all groups are visible after extraction
- Style the expander header with a subtle background color (e.g., `#F0F4F8`) to distinguish it from data rows

### How to verify
- [ ] Rows are grouped by layout name with collapsible headers (AC-02)
- [ ] Group headers show PR No, Project Name, Product Name, Spec, Surface Treatment, Color (AC-02)

---

## TASK-004-02-03: Implement summary panel using UniformGrid with 5 cells and colored backgrounds

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a summary bar section between the extraction panel and the DataGrid
- Use a `UniformGrid Columns="5"` containing 5 `Border` elements, each with a colored background:
  - Cell 1: "Total" with neutral background (`#E5E7EB`), bound to `{Binding TotalItems}`
  - Cell 2: "Valid" with green background (`#D1FAE5`), bound to `{Binding ValidItems}`
  - Cell 3: "Warnings" with yellow background (`#FEF3C7`), bound to `{Binding WarningItems}`
  - Cell 4: "Errors" with red background (`#FEE2E2`), bound to `{Binding ErrorItems}`
  - Cell 5: "Skipped" with gray background (`#F3F4F6`), bound to `{Binding SkippedItems}`
- Each cell contains a bold count `TextBlock` and a descriptive sub-label `TextBlock`
- Add a "Layouts: {LayoutsFound}" label above or next to the summary bar

### How to verify
- [ ] Summary panel displays all 5 metric cells with appropriate colors (AC-03)
- [ ] Total Layouts count is shown (AC-03)

---

## TASK-004-02-04: Create ValidationStatusToIconConverter for Status column DataTemplate

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Converters/ValidationStatusToIconConverter.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-02-01 |

### What to do
- Create `ValidationStatusToIconConverter` class implementing `IValueConverter` in namespace `OdooAutoCAD.App.Converters`
- Map `RowValidationStatus.Valid` to green checkmark symbol (Unicode `\u2714` or a `Path` geometry)
- Map `RowValidationStatus.Warning` to yellow triangle symbol (Unicode `\u26A0`)
- Map `RowValidationStatus.Error` to red exclamation symbol (Unicode `\u2757`)
- Map `RowValidationStatus.Pending` to gray dash or empty string
- Register the converter in `BOQManagerPage.xaml` resources: `<converters:ValidationStatusToIconConverter x:Key="ValidationStatusToIconConverter"/>`
- Use the converter in the Status column `DataTemplate`: `{Binding ValidationStatus, Converter={StaticResource ValidationStatusToIconConverter}}`

### How to verify
- [ ] Valid rows show green checkmark, warnings show yellow triangle, errors show red exclamation (AC-04)

---

## TASK-004-02-05: Bind AllEntries ObservableCollection to DataGrid ItemsSource with grouping

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-02-01 |
| Blocks | None |

### What to do
- Ensure `AllEntries` is declared as `ObservableCollection<BOQEntryRow>` with a public getter in `BOQManagerViewModel`
- After extraction completes in `ExecuteExtractAsync()`, clear `AllEntries` and repopulate it from the extracted layout data
- Also populate `LayoutGroups` (`ObservableCollection<BOQLayoutGroup>`) with the layout-level metadata for group header display
- Flatten all `BOQLayoutGroup.Entries` into `AllEntries` so the DataGrid can display all rows with grouping via `CollectionViewSource`
- Call `OnPropertyChanged(nameof(AllEntries))` after population to trigger UI update

### How to verify
- [ ] DataGrid shows all extracted entries after extraction completes (AC-01)
- [ ] Data is displayed exactly as extracted after MText unformatting (AC-06)

---

## TASK-004-02-06: Implement RowValidationStatusToBackgroundConverter for row-level coloring

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Converters/RowValidationStatusToBackgroundConverter.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Create `RowValidationStatusToBackgroundConverter` class implementing `IValueConverter` in namespace `OdooAutoCAD.App.Converters`
- Map `RowValidationStatus.Valid` to a light green `SolidColorBrush` (`#F0FFF4` or transparent/default)
- Map `RowValidationStatus.Warning` to a light yellow `SolidColorBrush` (`#FFFBEB`)
- Map `RowValidationStatus.Error` to a light red `SolidColorBrush` (`#FFF5F5`)
- Map `RowValidationStatus.Pending` to transparent (default background)
- Apply the converter to the `DataGrid.RowStyle` using a `Style.Setter` with `Property="Background"` and `Value="{Binding ValidationStatus, Converter={StaticResource RowValidationStatusToBackgroundConverter}}"`

### How to verify
- [ ] Error rows have red background, warning rows have yellow background, valid rows have default/green (AC-04)

---

## TASK-004-02-07: Wire summary count properties with OnPropertyChanged notifications

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Declare `[ObservableProperty]` fields in `BOQManagerViewModel`: `int _totalItems`, `int _validItems`, `int _warningItems`, `int _errorItems`, `int _skippedItems`, `int _layoutsFound`
- Create a private method `UpdateSummaryCounts()` that iterates `AllEntries` and tallies counts by `ValidationStatus`:
  - `TotalItems = AllEntries.Count`
  - `ValidItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Valid)`
  - `WarningItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Warning)`
  - `ErrorItems = AllEntries.Count(e => e.ValidationStatus == RowValidationStatus.Error)`
  - `SkippedItems` from the extraction skip counter
- Call `UpdateSummaryCounts()` at the end of `ExecuteExtractAsync()` and at the end of `ExecuteValidateAsync()`
- Property change notifications are auto-generated by `[ObservableProperty]` via source generators

### How to verify
- [ ] Summary counts update automatically after extraction (AC-05)
- [ ] Summary counts update automatically after validation (AC-05)

---

## Dependency Graph
```
TASK-004-02-04 (StatusToIconConverter)
       │
       └──▶ TASK-004-02-01 (DataGrid XAML)
                   │
                   ├──▶ TASK-004-02-02 (GroupStyle Expander)
                   └──▶ TASK-004-02-05 (Bind AllEntries)

TASK-004-02-03 (Summary Panel XAML) ─── independent
TASK-004-02-06 (RowBackgroundConverter) ─── independent
TASK-004-02-07 (Summary Count Properties) ─── independent

Tasks 03, 04, 06, 07 can all begin in parallel.
Task 01 depends on 04 (converter needed for status column).
Tasks 02 and 05 depend on 01 (DataGrid must exist).
```
