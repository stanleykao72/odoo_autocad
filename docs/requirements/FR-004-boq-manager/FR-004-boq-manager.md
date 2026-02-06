# FR-004: BOQ Manager Page

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: Not Started
> **Priority**: P1

## 1. Overview

The BOQ (Bill of Quantities) Manager Page is the central hub for the three-step BOQ workflow: **Extract** drawing data from AutoCAD tables, **Validate** the entries against Odoo product catalog and business rules, and **Push** the validated BOQ to Odoo via the `job_working_plan_boq.import2boq_v2` REST API. After a successful push, Odoo returns generated IDs which the system writes back into the AutoCAD drawing tables to maintain bidirectional traceability.

The Python reference implementation orchestrates this through `UtilPushToBoq.push_to_boq()`, which calls `autocad_util.get_layouts_values()` to extract layout and table data from all non-Model layouts, `odoo_util.import2boq(layout_dict)` to push the data to Odoo, and `autocad_util.set_layouts_tables_id(boq_list)` to write the returned `header_id` and `detail_id` values back into the AutoCAD table cells. The C# implementation must replicate this exact three-step workflow while adding a user-facing data grid for review, inline validation, and product mapping resolution before the push step.

Each AutoCAD layout contains attribute blocks (carrying header-level metadata such as `pr_no`, `project_name`, `job_working_plan_name`, `product_name`, `product_catelog`, `spec`, `surface_treatment`, `operation_flow`, `color_name`, `color_no`) and one or more legal tables with 9 columns (`position`, `product_no`, `width`, `height`, `len`, `thickness`, `qty`, `desc`, `detail_id`). A table is considered legal only if it has exactly 9 columns and the cell at row 0, column 7 contains the string `HEADER_ID`. Rows where both `qty` and `product_no` are empty are skipped during extraction.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-004-01 | CAD Engineer | Extract BOQ data from all AutoCAD layouts with one click | I do not have to manually copy table data from each layout |
| US-004-02 | CAD Engineer | See extracted BOQ data in a structured grid before pushing | I can review and verify the data before it goes to Odoo |
| US-004-03 | CAD Engineer | Validate BOQ entries against Odoo product catalog | I know which products are recognized and which need mapping |
| US-004-04 | CAD Engineer | Push validated BOQ entries to Odoo | The bill of quantities is recorded in the ERP system |
| US-004-05 | CAD Engineer | Have returned Odoo IDs written back to AutoCAD tables | My drawings reflect the Odoo record references for traceability |
| US-004-06 | CAD Engineer | See which rows were skipped during extraction and why | I can fix issues in the drawing if needed |
| US-004-07 | Project Manager | View BOQ generation summary (total, processed, skipped) | I can gauge data quality at a glance |
| US-004-08 | CAD Engineer | Map unrecognized AutoCAD product names to Odoo products | The push succeeds even when naming conventions differ |
| US-004-09 | CAD Engineer | See validation errors with row-level detail | I can navigate directly to the problematic entry |
| US-004-10 | CAD Engineer | Clear existing IDs from all tables before a fresh push | I can reset and re-push without stale references |
| US-004-11 | System Admin | View product mapping rules and modify them | Consistent mapping is maintained across sessions |
| US-004-12 | CAD Engineer | See progress indication during extraction and push | I know the system is working and not frozen |

## 3. Python Reference

### Source Files
- `utility/util_push_to_boq.py` - BOQ push orchestration (3-step workflow)
- `utility/util_autocad.py` - AutoCAD COM interface for layout/table extraction and ID writeback
- `utility/util_odoo.py` - Odoo REST API client (`import2boq`, `get_project`, `get_product`)
- `forms/form_main_modern.py` - Sidebar buttons and push_to_boq trigger (lines 343-355)

### Key Functions

**UtilPushToBoq (util_push_to_boq.py)**
- `push_to_boq()` - Orchestrates the full workflow: extract -> push to Odoo -> write IDs back

**UtilAutoCAD (util_autocad.py)**
- `get_layouts_values()` - Iterates all non-Model layouts, extracts attribute block values and table data; returns `{ 'all': [layout_dict, ...] }`
- `get_doc_layouts()` - Returns all layout objects excluding "Model"
- `get_layout_table_block(blocks)` - Filters blocks for `AcDbTable` objects that pass `chk_legal_table()`
- `chk_legal_table(table)` - Validates table has 9 columns and `HEADER_ID` marker at cell (0, 7)
- `get_table_data(table_list)` - Extracts rows into detail dicts; skips rows where both `qty` and `product_no` are empty; returns `(header_id, detail_list)`
- `get_layout_attribute_blocks_value(layout, tag_list)` - Extracts attribute block values for given tags
- `set_layouts_tables_id(boq_list)` - Writes `header_id` to cell (0, 8) and matched `detail_id` to cell (i, 8) for each table row
- `clear_table_id(layout)` - Clears column 8 for all rows except row 1 in a single layout
- `clear_all_tables_id()` - Clears column 8 across all layouts
- `LM_UnFormat(s, mtx)` - Strips AutoCAD MText formatting codes from cell values
- `get_detail_id_index(detail_list, product_no)` - Looks up `detail_id` by matching `product_no`

**UtilOdoo (util_odoo.py)**
- `import2boq(layout_dict)` - Calls `job_working_plan_boq.import2boq_v2` with layout data; returns list with `header_id` and `detail` per layout, or error message
- `get_project(pr_no)` - Looks up project by PR number; returns project dict with `id`, `name`, `job_working_plan_id`, `job_working_plan_name`
- `get_product()` - Fetches products filtered by category ID 27 and active status

### Python Sidebar Buttons (form_main_modern.py)
- "📋 從 Odoo 獲取參數" (`get_parameters_from_odoo`) - Opens parameter form for Odoo data
- "📊 推送到 BOQ" (`push_to_boq`) - Triggers `push_to_boq_util.push_to_boq()`
- "🔄 轉移 BOQ 到 PR" (`transfer_boq_to_pr`) - Converts BOQ to Purchase Requisition

### Layout Dict Structure (as sent to Odoo)
```python
{
    'all': [
        {
            'layout_name': str,
            'pr_no': str,
            'project_name': str,
            'job_working_plan_name': str,
            'product_name': str,
            'product_catelog': str,
            'spec': str,
            'surface_treatment': str,
            'operation_flow': str,
            'color_name': str,
            'color_no': str,
            'header_id': str,
            'detail': [
                {
                    'position': str,
                    'product_no': str,
                    'width': str,
                    'height': str,
                    'len': str,
                    'thickness': str,
                    'qty': str,
                    'desc': str,
                    'detail_id': str
                }
            ]
        }
    ]
}
```

### Table Column Structure (9 columns, 0-indexed)
| Index | Column Name | Description |
|-------|------------|-------------|
| 0 | position | Row position number |
| 1 | product_no | Product number/code |
| 2 | width | Width dimension |
| 3 | height | Height dimension |
| 4 | length | Length dimension |
| 5 | thickness | Thickness dimension |
| 6 | qty | Quantity |
| 7 | desc / HEADER_ID | Description (data rows) or "HEADER_ID" marker (header row 0) |
| 8 | detail_id | Odoo detail ID (written back after push) |

## 4. Functional Requirements

### Extraction

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-004-001 | The page SHALL provide an "Extract from AutoCAD" button that triggers extraction of layout and table data from all non-Model layouts in the active AutoCAD document | Must |
| FR-004-002 | Extraction SHALL iterate all document layouts (excluding "Model") and collect attribute block values for the tag list: `pr_no`, `project_name`, `job_working_plan_name`, `product_name`, `product_catelog`, `spec`, `surface_treatment`, `operation_flow`, `color_name`, `color_no` | Must |
| FR-004-003 | Extraction SHALL identify legal tables by verifying exactly 9 columns and the presence of `HEADER_ID` string at cell position (row 0, column 7) | Must |
| FR-004-004 | Extraction SHALL skip table rows where both `qty` (column 6) and `product_no` (column 1) are empty or whitespace-only after MText formatting removal | Must |
| FR-004-005 | Extraction SHALL apply MText unformatting (`LM_UnFormat` equivalent) to all cell values to strip AutoCAD formatting codes before populating the data grid | Must |
| FR-004-006 | Extraction SHALL read the `header_id` from cell (row 0, column 8) for each legal table | Must |
| FR-004-007 | The page SHALL display a progress indicator during extraction showing the current layout being processed | Should |

### Display

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-004-008 | The page SHALL display extracted BOQ data in a DataGrid with columns: Layout Name, Position, Product No, Width, Height, Length, Thickness, Qty, Description, Detail ID, and a Validation Status indicator | Must |
| FR-004-009 | The DataGrid SHALL group rows by layout name, with each layout group showing its header metadata (PR No, Project Name, Job Working Plan Name, Product Name, Spec, Surface Treatment, Color) in a collapsible group header | Should |
| FR-004-010 | The page SHALL display a summary panel showing: Total Layouts processed, Total Items extracted, Valid Items count, Invalid Items count, and Skipped Items count | Must |
| FR-004-011 | The DataGrid SHALL visually distinguish rows with validation errors (red background), warnings (yellow background), and valid rows (default/green background) | Should |

### Validation

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-004-012 | The page SHALL provide a "Validate" button that checks all extracted BOQ entries against Odoo product catalog and business rules before pushing | Must |
| FR-004-013 | Validation SHALL verify that each entry has a non-empty `product_no` that can be mapped to a valid Odoo product ID | Must |
| FR-004-014 | Validation SHALL verify that each entry has a positive numeric `qty` value | Must |
| FR-004-015 | Validation SHALL verify that a valid project context exists (non-zero `ProjectId` derived from `pr_no` lookup via `get_project`) | Must |
| FR-004-016 | Validation SHALL produce per-entry error details including entry index, field name, error message, and severity level (Error or Warning) using the `BOQValidationResult`/`BOQValidationError` model | Must |
| FR-004-017 | Validation results SHALL be displayed in a validation results panel below the DataGrid, with clickable errors that select/highlight the corresponding row in the grid | Should |

### Push to Odoo

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-004-018 | The page SHALL provide a "Push to Odoo" button that sends the validated layout data to Odoo via `IOdooService.ImportToBOQAsync()` (equivalent to the `import2boq_v2` API call) | Must |
| FR-004-019 | The Push operation SHALL be disabled until validation passes with zero Error-severity issues | Must |
| FR-004-020 | After a successful push, the system SHALL write the returned `header_id` and `detail_id` values back into the AutoCAD table cells (column 8) via `IAutoCADService.SetTableValue()`, matching `detail_id` to rows by `product_no` | Must |
| FR-004-021 | The Push operation SHALL display a progress indicator showing records processed out of total | Should |
| FR-004-022 | Upon push completion, the page SHALL display a result summary: records created, records updated, records failed, with any error messages from the Odoo response | Must |

### Product Mapping

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-004-023 | The page SHALL provide a Product Mapping panel or dialog where users can map unrecognized AutoCAD `product_no` values to Odoo product IDs | Must |
| FR-004-024 | Product mapping SHALL first check the local mapping cache (`IBOQProcessor.GetProductMappings()`) before querying Odoo via `IOdooService.SearchProductsAsync()` | Must |
| FR-004-025 | Product mapping SHALL support case-insensitive matching as implemented by the `StringComparer.OrdinalIgnoreCase` dictionary in `BOQProcessor` | Must |
| FR-004-026 | The Product Mapping dialog SHALL provide a search field that queries Odoo products and displays results for the user to select from | Should |
| FR-004-027 | User-defined product mappings SHALL persist via `IBOQProcessor.SetProductMapping()` and be available across extraction sessions within the same application lifecycle | Must |

### ID Writeback and Reset

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-004-028 | The page SHALL provide a "Clear All IDs" button that clears column 8 across all layouts and tables in the active drawing (equivalent to `clear_all_tables_id()`) | Should |
| FR-004-029 | ID writeback SHALL set `header_id` at cell (row 0, column 8) and `detail_id` at cell (row i, column 8) for each data row (i > 1), matching by `product_no` using `LM_UnFormat` to normalize cell text | Must |

## 5. UI Wireframe Description

```
+------------------------------------------------------------------+
|                       BOQ Manager Page                            |
+------------------------------------------------------------------+
|                                                                    |
|  Extraction Panel                                                  |
|  +--------------------------------------------------------------+ |
|  | Project: PR-2025-001 - Steel Frame Phase 2                   | |
|  | Document: structure_drawing_v3.dwg                            | |
|  | Layouts Found: 5 (excluding Model)                            | |
|  |                                                                | |
|  | [Extract from AutoCAD]  [Clear All IDs]  [Refresh]            | |
|  | Progress: [====================] Layout 3/5 extracting...     | |
|  +--------------------------------------------------------------+ |
|                                                                    |
|  Summary Bar                                                       |
|  +--------------------------------------------------------------+ |
|  | Total: 47  |  Valid: 42  |  Warnings: 3  |  Errors: 2  |     | |
|  | Skipped: 8 (empty rows)                                       | |
|  +--------------------------------------------------------------+ |
|                                                                    |
|  BOQ Data Grid                                                     |
|  +--------------------------------------------------------------+ |
|  | [v] Layout: Sheet-1 (PR: PR-2025-001, Product: Beam A,       | |
|  |     Spec: Q345B, Surface: Hot-Dip Galvanized)                | |
|  +------+----------+-----+------+-----+-----+----+--------+----+|
|  | Pos  | Prod No  | W   | H    | Len | Thk | Qty| Desc   | St ||
|  +------+----------+-----+------+-----+-----+----+--------+----+|
|  |  1   | BM-001   | 200 | 300  | 6000| 12  | 4  | Main   | OK ||
|  |  2   | BM-002   | 150 | 250  | 4500| 10  | 8  | Second | OK ||
|  |  3   | XX-999   | 100 | 100  | 3000| 8   | 2  | Custom | !! ||
|  +------+----------+-----+------+-----+-----+----+--------+----+|
|  | [v] Layout: Sheet-2 (PR: PR-2025-001, Product: Column B)     | |
|  +------+----------+-----+------+-----+-----+----+--------+----+|
|  |  1   | CL-001   | 300 | 300  | 3500| 14  | 6  | Corner | OK ||
|  |  ...                                                          | |
|  +--------------------------------------------------------------+ |
|                                                                    |
|  Validation Results Panel                                          |
|  +--------------------------------------------------------------+ |
|  | [!] Row 3 (Sheet-1): Product "XX-999" not found in Odoo      | |
|  |     -> [Map Product] [Ignore]                                 | |
|  | [!] Row 7 (Sheet-3): Qty is zero or negative                 | |
|  |     -> [Edit] [Remove]                                        | |
|  +--------------------------------------------------------------+ |
|                                                                    |
|  Push Controls                                                     |
|  +--------------------------------------------------------------+ |
|  | [Validate All]          [Push to Odoo]         [Cancel]       | |
|  |                                                                | |
|  | Push Status: Ready (42 valid entries, 2 errors to resolve)    | |
|  | Last Push: 2026-02-05 16:30 - 40 records created, 0 failed   | |
|  +--------------------------------------------------------------+ |
|                                                                    |
|  Product Mapping Dialog (modal, opened from Validation Results)    |
|  +--------------------------------------------------------------+ |
|  | AutoCAD Name: XX-999                                          | |
|  | Search Odoo: [_______________] [Search]                       | |
|  |                                                                | |
|  | Results:                                                       | |
|  | ( ) Custom Beam 999 - CB999 - Structural Steel               | |
|  | ( ) Custom Part X - CPX-01 - Miscellaneous                   | |
|  |                                                                | |
|  | [Apply Mapping]  [Cancel]                                      | |
|  +--------------------------------------------------------------+ |
+------------------------------------------------------------------+
```

### Layout Details
- **Extraction Panel**: Top section with project context, document info, and action buttons. Uses WPF `StackPanel` with horizontal button bar.
- **Summary Bar**: Horizontal strip with key metrics using `UniformGrid` of 5 cells with bold counts and colored backgrounds (green for valid, yellow for warnings, red for errors, gray for skipped).
- **BOQ Data Grid**: WPF `DataGrid` with `GroupStyle` for layout grouping. Each group header is collapsible and shows layout-level metadata. The "St" (Status) column uses a `DataTemplate` with colored icon (green checkmark for OK, red exclamation for error, yellow triangle for warning).
- **Validation Results Panel**: `ListView` of `BOQValidationError` entries with clickable items. Each item shows severity icon, row reference, and error message, plus inline action buttons.
- **Push Controls**: Bottom bar with Validate, Push, and Cancel buttons. Push button is bound to `CanExecute` that checks validation state. Includes status text showing readiness and last push result.
- **Product Mapping Dialog**: Modal `Window` opened from validation results. Contains search box, Odoo product result list with radio selection, and Apply/Cancel buttons.

## 6. Data Model

### ViewModel

```csharp
public class BOQManagerViewModel : ObservableObject
{
    // --- Project Context ---
    public string CurrentPRNo { get; set; }
    public string CurrentProjectName { get; set; }
    public string CurrentDocumentName { get; set; }
    public int ProjectId { get; set; }
    public string JobWorkingPlanName { get; set; }

    // --- Extraction State ---
    public bool IsExtracting { get; set; }
    public string ExtractionProgress { get; set; }        // e.g. "Layout 3/5"
    public double ExtractionProgressPercent { get; set; }  // 0.0 - 1.0
    public int LayoutsFound { get; set; }

    // --- BOQ Data ---
    public ObservableCollection<BOQLayoutGroup> LayoutGroups { get; set; } = new();
    public ObservableCollection<BOQEntryRow> AllEntries { get; set; } = new();

    // --- Summary Counts ---
    public int TotalItems { get; set; }
    public int ValidItems { get; set; }
    public int WarningItems { get; set; }
    public int ErrorItems { get; set; }
    public int SkippedItems { get; set; }

    // --- Validation ---
    public bool IsValidating { get; set; }
    public bool HasValidationErrors { get; set; }
    public ObservableCollection<BOQValidationErrorDisplay> ValidationErrors { get; set; } = new();

    // --- Push State ---
    public bool IsPushing { get; set; }
    public bool CanPush { get; set; }           // true when validated with 0 errors
    public string PushStatusText { get; set; }
    public string LastPushResult { get; set; }
    public DateTime? LastPushTime { get; set; }

    // --- Product Mapping ---
    public ObservableCollection<ProductMappingEntry> ProductMappings { get; set; } = new();

    // --- Commands ---
    public IAsyncRelayCommand ExtractCommand { get; }
    public IAsyncRelayCommand ValidateCommand { get; }
    public IAsyncRelayCommand PushToOdooCommand { get; }
    public IAsyncRelayCommand ClearAllIdsCommand { get; }
    public IRelayCommand RefreshCommand { get; }
    public IRelayCommand<BOQValidationErrorDisplay> NavigateToErrorCommand { get; }
    public IRelayCommand<BOQEntryRow> OpenProductMappingCommand { get; }
}

/// <summary>
/// Groups BOQ entries by layout, mirroring the Python layout_dict structure.
/// </summary>
public class BOQLayoutGroup
{
    public string LayoutName { get; set; } = string.Empty;
    public string PRNo { get; set; } = string.Empty;
    public string ProjectName { get; set; } = string.Empty;
    public string JobWorkingPlanName { get; set; } = string.Empty;
    public string ProductName { get; set; } = string.Empty;
    public string ProductCatalog { get; set; } = string.Empty;
    public string Spec { get; set; } = string.Empty;
    public string SurfaceTreatment { get; set; } = string.Empty;
    public string OperationFlow { get; set; } = string.Empty;
    public string ColorName { get; set; } = string.Empty;
    public string ColorNo { get; set; } = string.Empty;
    public string HeaderId { get; set; } = string.Empty;
    public ObservableCollection<BOQEntryRow> Entries { get; set; } = new();
}

/// <summary>
/// A single row in the BOQ data grid, corresponding to one table detail row.
/// </summary>
public class BOQEntryRow : ObservableObject
{
    public string LayoutName { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public string ProductNo { get; set; } = string.Empty;
    public string Width { get; set; } = string.Empty;
    public string Height { get; set; } = string.Empty;
    public string Length { get; set; } = string.Empty;
    public string Thickness { get; set; } = string.Empty;
    public string Qty { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string DetailId { get; set; } = string.Empty;

    // Validation state
    public RowValidationStatus ValidationStatus { get; set; } = RowValidationStatus.Pending;
    public string ValidationMessage { get; set; } = string.Empty;

    // Product mapping state
    public int? MappedOdooProductId { get; set; }
    public string MappedOdooProductName { get; set; } = string.Empty;
    public bool IsProductMapped { get; set; }
}

public enum RowValidationStatus
{
    Pending,
    Valid,
    Warning,
    Error
}

/// <summary>
/// Display model for validation errors in the results panel.
/// </summary>
public class BOQValidationErrorDisplay
{
    public int EntryIndex { get; set; }
    public string LayoutName { get; set; } = string.Empty;
    public string RowReference { get; set; } = string.Empty;   // e.g. "Row 3 (Sheet-1)"
    public string Field { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string Severity { get; set; } = "Error";
}

/// <summary>
/// Product mapping cache entry for display.
/// </summary>
public class ProductMappingEntry
{
    public string AutoCADName { get; set; } = string.Empty;
    public int OdooProductId { get; set; }
    public string OdooProductName { get; set; } = string.Empty;
}
```

## 7. API/Service Dependencies

| Service | Interface | Usage |
|---------|-----------|-------|
| AutoCAD Service | `IAutoCADService` | `GetLayoutsValues()` for extraction; `SetTableValue()` for ID writeback; `GetTables()`, `GetTableValue()` for table inspection; `GetLayouts()` for layout enumeration; `IsConnected` for pre-check |
| Odoo Service | `IOdooService` | `ImportToBOQAsync(entries)` for BOQ push; `GetProductAsync(id)` and `SearchProductsAsync(term)` for product validation/mapping; `GetProjectAsync(id)` for project context; `GetProductsAsync()` for catalog cache |
| BOQ Processor | `IBOQProcessor` | `GenerateBOQAsync(layoutData, options)` for BOQ generation; `ValidateBOQAsync(entries)` for validation; `PushToOdooAsync(entries)` for push with pre-validation; `MapProductAsync(name)` for product resolution; `GetProductMappings()` / `SetProductMapping()` for mapping management |
| GUI Proxy | `IGUIProxy` | `ExecuteInGuiAsync("extract_autocad_parameters")` to run AutoCAD COM extraction on STA thread; `ExecuteInGuiAsync()` for ID writeback operations; `ProcessRequests()` called by DispatcherTimer at 100ms interval |
| Navigation Service | `INavigationService` | `NavigateTo()` for page transitions from Dashboard quick actions |
| Logger | `ILogger<BOQManagerViewModel>` | Structured logging for extraction, validation, push, and error events |

## 8. Validation Rules

### Table Structure Validation

| Rule | Description |
|------|-------------|
| VR-004-001 | A table is legal only if `table.Columns == 9` (matching the Python `chk_legal_table` check) |
| VR-004-002 | A table is legal only if `table.GetCellValue(0, 7) == "HEADER_ID"` (header marker validation) |
| VR-004-003 | Tables that fail either check SHALL be silently skipped and counted in `SkippedItems` |
| VR-004-004 | Layouts named "Model" SHALL always be excluded from extraction |

### Row Completeness Validation

| Rule | Description |
|------|-------------|
| VR-004-005 | Rows where both `qty` (column 6) and `product_no` (column 1) are empty/whitespace after MText unformatting SHALL be skipped |
| VR-004-006 | A row with non-empty `product_no` but empty `qty` SHALL produce a Warning-severity validation error |
| VR-004-007 | A row with `qty` that is non-numeric or negative SHALL produce an Error-severity validation error on field "Quantity" |
| VR-004-008 | A row with `qty` equal to zero SHALL produce a Warning-severity validation error |

### Product Mapping Validation

| Rule | Description |
|------|-------------|
| VR-004-009 | Each `product_no` value SHALL be resolved against `IBOQProcessor.GetProductMappings()` (case-insensitive) first, then via `IOdooService.SearchProductsAsync()` |
| VR-004-010 | If a `product_no` cannot be resolved to an Odoo product and `BOQGenerationOptions.ValidateProducts` is true, an Error-severity validation error SHALL be produced on field "ProductId" |
| VR-004-011 | If `BOQGenerationOptions.AutoCreateProducts` is true, unresolved products SHALL produce a Warning instead of an Error |

### Project Context Validation

| Rule | Description |
|------|-------------|
| VR-004-012 | A valid `ProjectId` (> 0) must be available before push; if not, the push SHALL be blocked with an Error-severity validation message |
| VR-004-013 | The `pr_no` extracted from the attribute block SHALL match a project in Odoo; mismatch produces an Error-severity validation error |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| AutoCAD not connected | "AutoCAD is not connected. Please connect to AutoCAD before extracting BOQ data." | Disable Extract button; show connection status with link to Dashboard |
| No active document | "No active AutoCAD document found. Please open a drawing file." | Show in extraction panel status area |
| No legal tables found | "No valid BOQ tables found in the drawing. Tables must have exactly 9 columns with HEADER_ID marker." | Show warning in summary bar with zero items |
| COM operation timeout | "AutoCAD operation timed out. The drawing may be too large or AutoCAD is unresponsive." | Show error dialog with retry option; log timeout via `IGUIProxy` stats |
| COM thread conflict | "A thread conflict occurred during AutoCAD access. The operation will be retried via GUI proxy." | Automatically retry via `IGUIProxy.ExecuteInGuiAsync()`; log the conflict |
| Odoo not connected | "Odoo is not connected. Please connect to Odoo before pushing BOQ data." | Disable Push button; show connection status |
| Odoo API returns error_code | "Odoo returned an error: {error_message}" | Display the `error_message` from the response in the push result panel; do not proceed with ID writeback |
| Partial push failure | "Push completed with errors: {RecordsCreated} created, {RecordsFailed} failed. See details." | Show per-record errors in the push result summary; allow retry for failed records |
| Product mapping not found | "Product '{product_no}' could not be matched to an Odoo product." | Highlight row in DataGrid; show "Map Product" action in validation results panel |
| ID writeback failure | "Failed to write IDs back to AutoCAD table in layout '{layout_name}': {error}" | Show error for affected layout; already-pushed data remains in Odoo; user can retry writeback |
| Network timeout during push | "Connection to Odoo timed out during BOQ push. {ProcessedItems}/{TotalItems} items were processed." | Show partial completion status; allow retry for remaining items |
| Validation finds zero valid entries | "No valid entries to push. Please resolve all validation errors first." | Keep Push button disabled; highlight validation results |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/BOQManagerPage.xaml` - WPF page with DataGrid, validation panel, and push controls
- `Views/Pages/BOQManagerPage.xaml.cs` - Code-behind (minimal, for DataGrid interaction events)
- `ViewModels/BOQManagerViewModel.cs` - MVVM ViewModel with async commands and observable properties
- `Views/Dialogs/ProductMappingDialog.xaml` - Modal dialog for product mapping resolution
- `ViewModels/Dialogs/ProductMappingDialogViewModel.cs` - ViewModel for product mapping dialog

### Three-Step Workflow Implementation

The core BOQ workflow must preserve the exact sequence from the Python reference:

1. **Extract** (`autocad_util.get_layouts_values()` equivalent):
   - Execute via `IGUIProxy.ExecuteInGuiAsync("extract_layouts_values")` to ensure STA thread safety for COM operations.
   - The registered handler on the GUI thread calls `IAutoCADService.GetLayoutsValues()`.
   - Build `BOQLayoutGroup` and `BOQEntryRow` collections from the returned `LayoutData`.

2. **Push** (`odoo_util.import2boq(layout_dict)` equivalent):
   - Call `IBOQProcessor.PushToOdooAsync(entries)` which internally validates then calls `IOdooService.ImportToBOQAsync(entries)`.
   - The Odoo API endpoint `job_working_plan_boq.import2boq_v2` returns a list with `header_id` and `detail` items per layout.
   - Parse the response to build the writeback list.

3. **ID Writeback** (`autocad_util.set_layouts_tables_id(boq_list)` equivalent):
   - Execute via `IGUIProxy.ExecuteInGuiAsync("set_layouts_tables_id", parameters)` for COM thread safety.
   - For each layout, locate the legal tables, write `header_id` to cell (0, 8), then iterate data rows (i > 1) and match `product_no` (after MText unformatting) to find the corresponding `detail_id` from the Odoo response, writing it to cell (i, 8).

### Thread Safety for COM Operations

All AutoCAD COM interactions (extraction, ID writeback, table clearing) MUST be routed through `IGUIProxy` because:
- AutoCAD COM objects are STA (Single-Threaded Apartment) and must be accessed from the GUI main thread.
- The ViewModel executes on background threads via `IAsyncRelayCommand`.
- The `IGUIProxy` message queue architecture ensures requests are serialized and executed on the GUI thread's `DispatcherTimer` (100ms polling interval).
- The `BOQProcessor` already demonstrates this pattern: `_guiProxy.ExecuteInGuiAsync("extract_autocad_parameters")` in `GenerateBOQForProjectAsync()`.

### MVVM Bindings
- Use `CommunityToolkit.Mvvm` for `ObservableObject`, `AsyncRelayCommand`, `RelayCommand`.
- DataGrid `ItemsSource` bound to `AllEntries` with `CollectionViewSource` grouping by `LayoutName`.
- Validation status column uses `DataTemplate` with `{Binding ValidationStatus, Converter={StaticResource ValidationStatusToIconConverter}}`.
- Push button `Command="{Binding PushToOdooCommand}"` with `CanExecute` returning `!HasValidationErrors && !IsPushing && TotalItems > 0`.
- Summary counts update via property change notifications after extraction and validation.
- Progress bar `Value="{Binding ExtractionProgressPercent}"` and `Visibility="{Binding IsExtracting, Converter={StaticResource BoolToVisibilityConverter}}"`.

### MText Formatting Removal
The Python `LM_UnFormat()` method strips AutoCAD MText formatting codes (font changes, paragraph marks, special characters). The C# equivalent must apply the same regex-based substitution chain:
1. Replace `\\\\` with ASCII 032 placeholder.
2. Replace `\\P`, `\n`, `\t` with space.
3. Strip `\A`, `\C`, `\c`, `\F`, `\f`, `\H`, `\L`, `\l`, `\O`, `\o`, `\p`, `\Q`, `\T`, `\W` formatting commands and their arguments.
4. Strip `\S` stacking expressions.
5. Remove remaining escape characters and braces.

This should be implemented as a static utility method in a shared `MTextFormatter` class within `OdooAutoCAD.Core`.

### Differences from Python
- Python `push_to_boq()` runs all three steps synchronously in sequence. C# separates them into distinct user-triggered actions (Extract, Validate, Push) with a review step in between.
- Python has no intermediate validation step; it pushes directly. C# adds explicit validation with `IBOQProcessor.ValidateBOQAsync()` before allowing push.
- Python stores no product mapping cache; it relies on Odoo to resolve products. C# adds a local mapping dictionary (`Dictionary<string, int>`) in `BOQProcessor` for user-defined overrides.
- Python sidebar buttons are flat actions. C# provides a dedicated page with rich data grid, summary, and validation panel.
- Python uses `self.log.safe_log_insert()` for all output. C# uses structured logging (`ILogger`) plus ViewModel observable properties for UI feedback.
