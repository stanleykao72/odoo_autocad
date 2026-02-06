# US-004-01: Extract BOQ Data

## User Story
**As a** CAD Engineer,
**I want to** extract BOQ data from all AutoCAD layouts with one click,
**So that** I do not have to manually copy table data from each layout.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Clicking "Extract from AutoCAD" iterates all non-Model layouts in the active document and collects attribute block values and table data
- [ ] AC-02: Only legal tables (exactly 9 columns with `HEADER_ID` at cell (0, 7)) are processed; others are silently skipped
- [ ] AC-03: Attribute block values are extracted for all required tags: `pr_no`, `project_name`, `job_working_plan_name`, `product_name`, `product_catelog`, `spec`, `surface_treatment`, `operation_flow`, `color_name`, `color_no`
- [ ] AC-04: Table rows where both `qty` (column 6) and `product_no` (column 1) are empty/whitespace after MText unformatting are skipped
- [ ] AC-05: MText formatting codes are stripped from all cell values using `MTextFormatter.UnFormat()` before populating the data grid
- [ ] AC-06: The `header_id` is read from cell (row 0, column 8) for each legal table
- [ ] AC-07: A progress indicator displays the current layout being processed (e.g., "Layout 3/5")
- [ ] AC-08: All AutoCAD COM operations are routed through `IGUIProxy.ExecuteInGuiAsync()` for STA thread safety
- [ ] AC-09: If AutoCAD is not connected, the Extract button is disabled and a message is shown: "AutoCAD is not connected. Please connect to AutoCAD before extracting BOQ data."
- [ ] AC-10: If no active document is found, the message "No active AutoCAD document found. Please open a drawing file." is shown

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-001 | Provide "Extract from AutoCAD" button that triggers extraction from all non-Model layouts | Must |
| FR-004-002 | Iterate all layouts (excluding Model) and collect attribute block values for the tag list | Must |
| FR-004-003 | Identify legal tables by verifying 9 columns and HEADER_ID at cell (0, 7) | Must |
| FR-004-004 | Skip rows where both qty and product_no are empty/whitespace after MText unformatting | Must |
| FR-004-005 | Apply MText unformatting to all cell values to strip AutoCAD formatting codes | Must |
| FR-004-006 | Read header_id from cell (row 0, column 8) for each legal table | Must |
| FR-004-007 | Display progress indicator during extraction showing current layout | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-01-01 | Add "Extract from AutoCAD" button to BOQ Manager page with binding to ExtractCommand | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-01-02 | Implement ExtractCommand in ViewModel that calls IGUIProxy for COM extraction | `ViewModels/BOQManagerViewModel.cs` | L |
| TASK-004-01-03 | Implement GetLayoutsValues() in IAutoCADService to iterate layouts, validate tables, and extract data | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | L |
| TASK-004-01-04 | Implement MTextFormatter.UnFormat() static method for stripping MText formatting codes | `OdooAutoCAD.Core/Utilities/MTextFormatter.cs` | M |
| TASK-004-01-05 | Register extraction handler in IGUIProxy for STA thread execution | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | M |
| TASK-004-01-06 | Add progress reporting with IProgress<T> pattern during layout iteration | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-01-07 | Implement chk_legal_table equivalent: validate 9 columns and HEADER_ID marker | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-01-08 | Implement row skip logic for empty qty + product_no rows | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |

## Dependencies
- Depends on: US-002-01 (AutoCAD connected)
- Blocks: US-004-02, US-004-03, US-004-06, US-004-07

## Notes
- The extraction must preserve the exact sequence from the Python reference: `autocad_util.get_layouts_values()` iterates all non-Model layouts, filters for legal tables via `chk_legal_table()`, extracts attribute blocks, and builds the layout dict structure.
- COM operations must be routed through `IGUIProxy` because AutoCAD COM objects are STA and must be accessed from the GUI main thread. The ViewModel executes on background threads via `IAsyncRelayCommand`.
- The MText unformatting regex chain must match the Python `LM_UnFormat()` behavior exactly: handle `\\\\`, `\\P`, `\n`, `\t`, and all formatting commands (`\A`, `\C`, `\F`, `\H`, `\L`, `\O`, `\p`, `\Q`, `\S`, `\T`, `\W`, etc.).
