# US-002-03: Extract Parameters

## User Story
**As a** CAD Engineer,
**I want to** extract parameters from a layout,
**So that** I can map drawing data to Odoo products.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Clicking "Extract Parameters" retrieves block attributes from the selected layout (pr_no, project_name, job_working_plan_name, product_name, product_catelog, spec, surface_treatment, operation_flow, color_name, color_no)
- [ ] AC-02: Extracted block attributes are displayed as key-value pairs in the Layout Details panel
- [ ] AC-03: Table data from valid layout tables is displayed in a DataGrid with columns: Position, Product No, Width, Height, Length, Thickness, Qty, Description
- [ ] AC-04: Table structure is validated to have exactly 9 columns and "HEADER_ID" in column 7 header; invalid tables are skipped with a warning logged
- [ ] AC-05: Empty rows (where both qty and product_no are empty) are filtered out from the displayed table data
- [ ] AC-06: AutoCAD MText formatting codes are stripped from all extracted values using an LM_UnFormat equivalent
- [ ] AC-07: The Detail ID column (index 8) is stored internally but not displayed to the user
- [ ] AC-08: Row 0 (header with header_id) and Row 1 (labels) are skipped; data extraction starts from Row 2
- [ ] AC-09: Extract Parameters button is only enabled when AutoCAD is connected and a layout is selected
- [ ] AC-10: All extraction COM operations execute on the GUI/STA thread via IGUIProxy

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-013 | Page SHALL display extracted PR Number from the active layout | Must |
| FR-002-014 | Page SHALL display Project Name, Job Working Plan Name from layout attributes | Must |
| FR-002-015 | Page SHALL show extracted block attributes (product_name, spec, category, etc.) | Must |
| FR-002-016 | Page SHALL display table data from valid layout tables in a DataGrid | Must |
| FR-002-017 | Table data SHALL show: position, product_no, width, height, len, thickness, qty, desc | Must |
| FR-002-018 | Page SHALL validate table structure (9 columns, HEADER_ID in col 7) | Must |
| FR-002-019 | Empty rows (both qty and product_no empty) SHALL be filtered out | Must |
| FR-002-020 | Text formatting codes SHALL be stripped via LM_UnFormat equivalent | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-03-01 | Add Layout Details panel (right panel) and Table Data DataGrid to XAML | `Views/Pages/AutoCADPage.xaml` | M |
| TASK-002-03-02 | Implement ExtractParametersCommand in ViewModel | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-002-03-03 | Implement GetLayoutsValues() and GetLayoutValues(name) in AutoCADService | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | L |
| TASK-002-03-04 | Implement get_attribute_values equivalent for block attribute extraction | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | L |
| TASK-002-03-05 | Implement get_table_data equivalent with 9-column validation and HEADER_ID check | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | L |
| TASK-002-03-06 | Implement LM_UnFormat equivalent in C# (regex to strip MText formatting codes: \P, \C, \F, \H, \S, brace formatting) | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-03-07 | Implement empty row filtering logic (skip rows where both qty and product_no are empty) | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | S |
| TASK-002-03-08 | Define TableRowData model class and LayoutAttributes dictionary in ViewModel | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-03-09 | Bind DataGrid columns to TableRowData properties (Position, ProductNo, Width, Height, Length, Thickness, Quantity, Description) | `Views/Pages/AutoCADPage.xaml` | M |

## Dependencies
- Depends on: US-002-01 (AutoCAD connection), US-002-02 (layout selection)
- Blocks: US-002-05 (PR/project info depends on extraction)

## Notes
- The 10 standard tag attributes extracted from layouts are: pr_no, project_name, job_working_plan_name, product_name, product_catelog, spec, surface_treatment, operation_flow, color_name, color_no.
- Table column mapping: Index 0=Position, 1=Product No, 2=Width, 3=Height, 4=Length, 5=Thickness, 6=Quantity, 7=Description, 8=Detail ID (hidden).
- The Python `LM_UnFormat(s, mtx)` function removes AutoCAD MText formatting codes. The C# equivalent should use regex to strip `\P`, `\C`, `\F`, `\H`, `\S`, and brace formatting sequences.
- Validation rules VR-002-003 (exactly 9 columns), VR-002-004 (HEADER_ID in col 7), and VR-002-005 (empty row filtering) all apply.
- Validation rule VR-002-007: Extract Parameters requires both AutoCAD connected and project context.
