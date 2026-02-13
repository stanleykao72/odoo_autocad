# US-010-03: File-Based Writeback

## User Story
**As a** CAD Engineer,
**I want to** write back IDs and attributes in file mode,
**So that** my drawing data stays synchronized with Odoo even without COM.

## Parent Feature
- **FR**: [FR-010-dual-mode-autocad](FR-010-dual-mode-autocad.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Block attribute writes produce a DXF file via ACadSharp DxfWriter
- [ ] AC-02: Table ID writeback creates a sidecar JSON file (`{name}.boq-ids.json`)
- [ ] AC-03: Sidecar JSON contains version, source_dwg, modified_at, layouts with header_id and details
- [ ] AC-04: Original DWG file is never modified or overwritten
- [ ] AC-05: `IDrawingDataService.SupportsWrite` returns true in file mode (for attributes via DXF)
- [ ] AC-06: `IDrawingDataService.RecommendedWriteStrategy` returns `SidecarJson` for table IDs
- [ ] AC-07: BOQ page shows info banner explaining sidecar writeback in file mode
- [ ] AC-08: Sidecar JSON can be loaded when switching back to COM mode for actual writeback
- [ ] AC-09: SidecarIdStore handles concurrent access safely

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-010-016 | Block attribute writes SHALL use ACadSharp DXF export | Must |
| FR-010-017 | Table ID writeback SHALL use sidecar JSON as primary strategy | Must |
| FR-010-018 | Sidecar JSON SHALL be saved alongside the DWG file | Must |
| FR-010-019 | Original DWG files SHALL never be overwritten | Must |
| FR-010-020 | `SupportsWrite` SHALL indicate write capability | Must |
| FR-010-021 | `RecommendedWriteStrategy` SHALL guide UI behavior | Must |
| FR-010-024 | BOQ page SHALL show info banner in file mode | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-010-03-01 | Implement SidecarIdStore — JSON sidecar read/write | `OdooAutoCAD.Core/AutoCAD/SidecarIdStore.cs` | M |
| TASK-010-03-02 | Implement FileDrawingDataService — write: attributes via DXF, table IDs via sidecar | `OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` | M |
| TASK-010-03-03 | Implement FileDrawingDataService.SaveAsDxf() for DXF export | `OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` | M |
| TASK-010-03-04 | Refactor BOQViewModel for sidecar writeback fallback when !SupportsWrite | `ViewModels/BOQViewModel.cs` | M |
| TASK-010-03-05 | Add file mode info banner to BOQ page | `Views/Pages/BOQPage.xaml` | S |
| TASK-010-03-06 | Tests: SidecarIdStore (roundtrip, concurrent access) | `tests/` | M |
| TASK-010-03-07 | Tests: FileDrawingDataService write operations | `tests/` | M |

## Dependencies
- Depends on: US-010-02 (file read infrastructure)
- Blocks: None (end-of-chain for file mode)

## Notes
- **Sidecar JSON** is the primary write strategy for table IDs because ACadSharp's TableEntity write support is marked WIP and may produce corrupted output.
- DXF export via `DxfWriter` is stable for block attribute modifications across all AutoCAD versions.
- The sidecar file uses the pattern `{dwg_basename}.boq-ids.json` and is placed in the same directory as the DWG.
- When switching from File mode back to COM mode, the application can optionally read the sidecar JSON and write the IDs into AutoCAD via COM — this is a convenience feature, not a requirement.
- `SaveAsync()` in file mode saves modified attributes as DXF and triggers sidecar JSON write.
