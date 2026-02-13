# US-010-02: File-Based Data Extraction

## User Story
**As a** CAD Engineer,
**I want to** extract BOQ data and parameters from a DWG file without COM,
**So that** I can process drawings from AutoCAD LT.

## Parent Feature
- **FR**: [FR-010-dual-mode-autocad](FR-010-dual-mode-autocad.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: User can load a DWG file via file picker in AutoCAD page (File mode)
- [ ] AC-02: Layouts are extracted from DWG file excluding "Model"
- [ ] AC-03: Table data (9-column format) is extracted from layout tables
- [ ] AC-04: Block attributes are extracted from attribute blocks in each layout
- [ ] AC-05: Header IDs are read from table cell(0,8) per layout
- [ ] AC-06: PR number is extracted from block reference text/attributes
- [ ] AC-07: Extracted data matches COM mode output for the same DWG file
- [ ] AC-08: File info (name, version, layout count) is displayed in connection panel
- [ ] AC-09: ACadSharp table read failures are caught and return empty data (no crash)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-010-010 | File mode SHALL read layouts from DWG via ACadSharp | Must |
| FR-010-011 | File mode SHALL extract table data (9-column tables) from DWG | Must |
| FR-010-012 | File mode SHALL extract block attributes from DWG | Must |
| FR-010-013 | File mode SHALL extract header IDs from table cells | Must |
| FR-010-014 | File mode SHALL extract PR number from block references | Must |
| FR-010-015 | File mode SHALL support loading DWG files via file picker | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-010-02-01 | Define IDwgFileService interface (extends IDwgReaderService + write + file info) | `OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs` | S |
| TASK-010-02-02 | Implement DwgFileService — file info, version detection, CanWriteDwg | `OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` | M |
| TASK-010-02-03 | Implement DwgFileService.ExtractTableData() — read TableEntity cells | `OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` | M |
| TASK-010-02-04 | Implement DwgFileService.GetHeaderIds() — read header IDs from tables | `OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` | S |
| TASK-010-02-05 | Implement FileDrawingDataService — read operations delegation | `OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` | M |
| TASK-010-02-06 | Add file picker and DWG info display to AutoCAD page | `Views/Pages/AutoCADPage.xaml` | M |
| TASK-010-02-07 | Refactor AutoCADViewModel to use IDrawingDataService for extraction | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-010-02-08 | Tests: DwgFileService (file info, table read, attribute extraction) | `tests/` | M |

## Dependencies
- Depends on: US-010-01 (mode infrastructure, IDrawingDataService interface)
- Depends on: ACadSharp submodule setup
- Blocks: US-010-03 (write operations need read infrastructure)

## Notes
- ACadSharp's `DwgReader` is the primary read path. `DxfReader` can be used as fallback for DXF files.
- Table extraction via ACadSharp TableEntity may be unstable — wrap in try-catch, return empty on failure.
- The existing `IDwgReaderService` + `DwgReaderService` already handle layout and attribute reads. `IDwgFileService` extends this with table data and file info.
- Block identification follows the same logic as COM mode: find AcDbBlockReference with `project_name` or `job_working_plan_name` attribute tag.
- Data rows start at row 2 (row 0=title with header_id, row 1=labels, row 2+=data).
