# TASKS: US-002-03 — Extract Parameters

> **Parent US**: [US-002-03](US-002-03-extract-parameters.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P1
> **Tasks**: 9 | **Effort**: 2S + 4M + 3L
> **Status**: Done

## Prerequisites
- [x] US-002-01 (AutoCAD connection must be established)
- [x] US-002-02 (layout selection must be available)

## Acceptance Criteria
- [x] AC-01: Clicking "Extract Parameters" retrieves block attributes from the selected layout (pr_no, project_name, job_working_plan_name, product_name, product_catelog, spec, surface_treatment, operation_flow, color_name, color_no)
- [x] AC-02: Extracted block attributes are displayed as key-value pairs in the Layout Details panel
- [x] AC-03: Table data from valid layout tables is displayed in a DataGrid with columns: Position, Product No, Width, Height, Length, Thickness, Qty, Description
- [x] AC-04: Table structure is validated to have exactly 9 columns and "HEADER_ID" in column 7 header; invalid tables are skipped with a warning logged
- [x] AC-05: Empty rows (where both qty and product_no are empty) are filtered out from the displayed table data
- [x] AC-06: AutoCAD MText formatting codes are stripped from all extracted values using an LM_UnFormat equivalent
- [x] AC-07: The Detail ID column (index 8) is stored internally but not displayed to the user
- [x] AC-08: Row 0 (header with header_id) and Row 1 (labels) are skipped; data extraction starts from Row 2
- [x] AC-09: Extract Parameters button is only enabled when AutoCAD is connected and a layout is selected
- [x] AC-10: All extraction COM operations execute on the GUI/STA thread via IGUIProxy

---

## TASK-002-03-01: Add Layout Details panel and Table Data DataGrid to XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-03-09 |

### What to do
- Add a right-side panel (Layout Details) in the two-column layout area of `AutoCADPage.xaml`:
  - Use an `ItemsControl` or `ListView` bound to `LayoutAttributes` (`Dictionary<string, string>`) to display key-value pairs
  - Each item shows tag name (left) and value (right) in a two-column grid or `UniformGrid`
  - Labels for the 10 standard attributes: PR No, Project Name, Job Working Plan, Product Name, Product Catalog, Spec, Surface Treatment, Operation Flow, Color Name, Color No
- Add a full-width Table Data section below the layout panels:
  - Use a `DataGrid` with `AutoGenerateColumns="False"` bound to `TableRows` (`ObservableCollection<TableRowData>`)
  - DataGrid should be read-only (`IsReadOnly="True"`)
- Add an Actions Bar at the bottom with an "Extract Parameters" `Button`:
  - Bind `Command` to `ExtractParametersCommand`
  - Bind `IsEnabled` to a multi-condition: `IsConnected && SelectedLayout != null` (VR-002-007)

### How to verify
- [x] Layout Details panel shows key-value attribute pairs (AC-02)
- [x] DataGrid is present for table data display (AC-03)
- [x] Extract Parameters button is present and disabled when conditions not met (AC-09)

---

## TASK-002-03-02: Implement ExtractParametersCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-002-03-08 |
| Blocks | None |

### What to do
- Add `[RelayCommand(CanExecute = nameof(CanExtractParameters))]` on an `async Task ExtractParametersAsync()` method
- Implement `bool CanExtractParameters()` returning `IsConnected && SelectedLayout != null`
- In `ExtractParametersAsync()`:
  - Call `_guiProxy.ExecuteInGuiAsync("extract_layout_values", new Dictionary<string, object?> { { "layoutName", SelectedLayout.Name } }, timeout: 30000)`
  - Parse the `GUIProxyResponse.Result` as `LayoutData`
  - Populate `LayoutAttributes` dictionary from `LayoutData.Parameters` (filtering to the 10 standard tags)
  - Populate `TableRows` from `LayoutData.Tables` using the table processing logic
  - Set `HeaderId` from table header row (row 0, column 8)
  - Update `PRNumber`, `ProjectName`, `JobWorkingPlanName` from extracted attributes
- Handle errors: display error message in `ErrorMessage` property
- Notify property changes for `LayoutAttributes` and `TableRows`

### How to verify
- [x] ExtractParametersCommand calls IGUIProxy for thread-safe extraction (AC-10)
- [x] Extracted attributes populate the LayoutAttributes dictionary (AC-01, AC-02)
- [x] Extract button only enabled when connected and layout selected (AC-09)

---

## TASK-002-03-03: Implement GetLayoutsValues() and GetLayoutValues(name) in AutoCADService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | L |
| Depends On | TASK-002-03-04, TASK-002-03-05 |
| Blocks | None |

### What to do
- Refactor the existing `GetLayoutsValues()` to iterate through the specified layout's `PaperSpace` (not just `ModelSpace`):
  - Switch to the target layout using `SwitchToLayout(layoutName)`
  - Get the layout's block reference: `_acadDoc.PaperSpace` for the active layout
  - Iterate entities in the layout space, not ModelSpace
- Refactor `GetLayoutValues(string layoutName)`:
  - Switch to the named layout
  - Call block attribute extraction (`ExtractBlockAttributes`) for all `AcDbBlockReference` entities
  - Call table data extraction (`ExtractTableData`) for all `AcDbTable` entities
  - Return a complete `LayoutData` with `Parameters` and `Tables` populated
- Register IGUIProxy handler:
  - `"extract_layout_values"` action: takes `layoutName` parameter, calls `GetLayoutValues(layoutName)`, returns `LayoutData`
- Ensure `LayoutData.LayoutName` is set to the layout name
- Ensure `LayoutData.ExtractedAt` is set to `DateTime.UtcNow`

### How to verify
- [x] GetLayoutValues extracts both block attributes and table data from the layout (AC-01, AC-03)
- [x] Extraction works on PaperSpace entities for the specified layout

---

## TASK-002-03-04: Implement get_attribute_values equivalent for block attribute extraction

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-002-03-03 |

### What to do
- Refactor or extend the existing `ExtractBlockAttributes(dynamic block, Dictionary<string, object> parameters)` method:
  - Accept a tag list filter: `string[] tagList` with the 10 standard tags: `pr_no`, `project_name`, `job_working_plan_name`, `product_name`, `product_catelog`, `spec`, `surface_treatment`, `operation_flow`, `color_name`, `color_no`
  - Iterate `block.GetAttributes()` and for each attribute:
    - Get `attr.TagString` (tag name) and `attr.TextString` (value)
    - If `tagList` is provided, only include attributes whose tag matches (case-insensitive)
    - Apply `LMUnFormat()` (TASK-002-03-06) to strip MText formatting from values
    - Add to parameters dictionary: `parameters[tag] = cleanedValue`
  - Handle blocks without attributes gracefully (`block.HasAttributes` check)
- Create a public method signature: `Dictionary<string, string> GetAttributeValues(string layoutName, string blockName, string[] tagList)`
- This method should be callable independently or as part of `GetLayoutValues`

### How to verify
- [x] Block attributes are extracted for the 10 standard tags (AC-01)
- [x] Values have MText formatting stripped (AC-06)

---

## TASK-002-03-05: Implement get_table_data equivalent with validation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | L |
| Depends On | TASK-002-03-06, TASK-002-03-07 |
| Blocks | TASK-002-03-03 |

### What to do
- Refactor the existing `ExtractTableData(dynamic table)` method to add validation and row processing:
  - **Column validation (VR-002-003)**: Check `table.Columns == 9`; if not 9 columns, skip table and log warning: "Table in layout has invalid structure (expected 9 columns)."
  - **Header validation (VR-002-004)**: Check that cell at row 0, column 7 contains "HEADER_ID"; if not, skip table and log warning
  - **Header ID extraction**: Extract `header_id` value from row 0, column 8 (`detail_id` column in header row)
  - **Row processing**: Skip row 0 (header) and row 1 (labels); start data extraction from row 2
  - For each data row (row 2+):
    - Map 9 columns to `TableRowData` properties: Position(0), ProductNo(1), Width(2), Height(3), Length(4), Thickness(5), Quantity(6), Description(7), DetailId(8)
    - Apply `LMUnFormat()` to all cell values
    - Apply empty row filtering (TASK-002-03-07): skip if both qty and product_no are empty
  - Return `List<TableRowData>` with processed rows
- Create a new method: `List<TableRowData> ProcessTableData(dynamic table, out string? headerId)`
- Ensure structured logging for skipped tables and filtered rows

### How to verify
- [x] Tables with != 9 columns are skipped with warning (AC-04)
- [x] HEADER_ID validation in column 7 header (AC-04)
- [x] Data starts from row 2, skipping rows 0 and 1 (AC-08)
- [x] Empty rows are filtered out (AC-05)

---

## TASK-002-03-06: Implement LM_UnFormat equivalent in C#

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-03-05 |

### What to do
- Create a static utility method in `AutoCADService` or a separate `MTextFormatter` utility class in namespace `OdooAutoCAD.Core.AutoCAD`:
  ```csharp
  public static string LMUnFormat(string input)
  ```
- Implement regex-based stripping of AutoCAD MText formatting codes:
  - `\P` -- paragraph break, replace with empty or newline
  - `\C\d+;` -- color codes (e.g., `\C1;`), remove entirely
  - `\F[^;]+;` -- font specifications, remove entirely
  - `\H\d+(\.\d+)?x?;` -- height codes, remove entirely
  - `\S[^;]+;` -- stacking/fraction codes, remove entirely
  - `\W\d+(\.\d+)?;` -- width factor, remove entirely
  - `\T\d+(\.\d+)?;` -- tracking factor, remove entirely
  - `\Q\d+;` -- obliquing angle, remove entirely
  - `\\\\` -- escaped backslash, replace with `\`
  - `\{` and `\}` -- literal braces, replace with `{` and `}`
  - Unmatched `{` and `}` -- grouping braces, remove
  - `\L`, `\l`, `\O`, `\o`, `\K`, `\k` -- underline/overline/strikethrough toggles, remove
- Handle null/empty input gracefully (return empty string)
- Add unit tests in `AutoCADServiceTests.cs` for each formatting code pattern

### How to verify
- [x] MText formatting codes are stripped from all values (AC-06)
- [x] Common patterns (\P, \C, \F, \H, \S) are correctly removed

---

## TASK-002-03-07: Implement empty row filtering logic

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-03-05 |

### What to do
- Create a filtering method used during table row processing:
  ```csharp
  private static bool IsEmptyRow(TableRowData row)
  {
      return string.IsNullOrWhiteSpace(row.Quantity) && string.IsNullOrWhiteSpace(row.ProductNo);
  }
  ```
- Apply this filter in `ProcessTableData` when building the result list:
  - After creating each `TableRowData` from a table row, check `IsEmptyRow()`
  - If true, skip the row (do not add to result list)
  - Log skipped rows at Debug level for troubleshooting
- VR-002-005: Both `qty` (column 6) AND `product_no` (column 1) must be empty for a row to be filtered out
- If only one is empty, the row should still be included

### How to verify
- [x] Rows where both qty and product_no are empty are excluded (AC-05)
- [x] Rows where only one is empty are retained

---

## TASK-002-03-08: Define TableRowData model class and LayoutAttributes in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-03-02 |

### What to do
- Create `TableRowData` class in namespace `OdooAutoCAD.Core.AutoCAD` (or in ViewModel namespace if preferred):
  ```csharp
  public class TableRowData
  {
      public string Position { get; set; } = string.Empty;
      public string ProductNo { get; set; } = string.Empty;
      public string Width { get; set; } = string.Empty;
      public string Height { get; set; } = string.Empty;
      public string Length { get; set; } = string.Empty;
      public string Thickness { get; set; } = string.Empty;
      public string Quantity { get; set; } = string.Empty;
      public string Description { get; set; } = string.Empty;
      public string DetailId { get; set; } = string.Empty;  // Internal, not displayed
  }
  ```
- Add to `AutoCADViewModel`:
  - `public ObservableCollection<TableRowData> TableRows { get; } = new();` -- table data rows
  - `[ObservableProperty] private Dictionary<string, string> _layoutAttributes = new();` -- extracted block attributes as key-value pairs
  - `[ObservableProperty] private string _headerId = string.Empty;` -- header ID from table row 0, column 8
- Ensure `TableRowData` is in a shared location accessible by both Core and App projects

### How to verify
- [x] TableRowData has all 9 properties matching table columns (AC-03, AC-07)
- [x] LayoutAttributes dictionary is available for key-value display (AC-02)

---

## TASK-002-03-09: Bind DataGrid columns to TableRowData properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | M |
| Depends On | TASK-002-03-01 |
| Blocks | None |

### What to do
- In the `DataGrid` in `AutoCADPage.xaml`, define explicit columns:
  ```xml
  <DataGrid.Columns>
      <DataGridTextColumn Header="Position" Binding="{Binding Position}" Width="80" />
      <DataGridTextColumn Header="Product No" Binding="{Binding ProductNo}" Width="120" />
      <DataGridTextColumn Header="Width" Binding="{Binding Width}" Width="80" />
      <DataGridTextColumn Header="Height" Binding="{Binding Height}" Width="80" />
      <DataGridTextColumn Header="Length" Binding="{Binding Length}" Width="80" />
      <DataGridTextColumn Header="Thickness" Binding="{Binding Thickness}" Width="80" />
      <DataGridTextColumn Header="Qty" Binding="{Binding Quantity}" Width="60" />
      <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*" />
  </DataGrid.Columns>
  ```
- Do NOT include a column for `DetailId` (AC-07: stored internally but not displayed)
- Set `ItemsSource="{Binding TableRows}"`
- Style the DataGrid with alternating row colors for readability
- Add a "No data" message overlay bound to `TableRows.Count == 0` visibility

### How to verify
- [x] DataGrid shows 8 visible columns matching the required fields (AC-03)
- [x] DetailId column is not displayed to the user (AC-07)
- [x] Column bindings correctly map to TableRowData properties

---

## Dependency Graph
```
TASK-002-03-07 (Empty row filtering)
    |
TASK-002-03-06 (LMUnFormat)
    |
    +---> TASK-002-03-05 (Table data with validation)
              |
TASK-002-03-04 (Block attribute extraction)
    |         |
    +---> TASK-002-03-03 (GetLayoutValues service)

TASK-002-03-08 (TableRowData model)
    |
    +---> TASK-002-03-02 (ExtractParametersCommand)

TASK-002-03-01 (XAML panels)
    |
    +---> TASK-002-03-09 (DataGrid column bindings)
```
