# US-011-01: DXF TABLE Entity Writer

## User Story
**As a** CAD Engineer,
**I want to** have file mode write header_id and detail_id directly into DXF TABLE cells,
**So that** I can open the exported DXF in AutoCAD and see the ID values in the table without relying on sidecar JSON.

## Parent Feature
- **FR**: [FR-011-table-entity-writer](FR-011-table-entity-writer.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: ACadSharp DxfWriter outputs TableEntity instead of skipping it
- [ ] AC-02: DXF TABLE output includes AcDbBlockReference and AcDbTable subclass markers
- [ ] AC-03: DXF TABLE preserves row count, column count, row heights, column widths
- [ ] AC-04: DXF TABLE preserves cell text content values
- [ ] AC-05: DXF TABLE preserves table style handle reference
- [ ] AC-06: DXF output is readable by AutoCAD 2010+ (valid DXF structure)
- [ ] AC-07: DwgFileService.WriteTableIdsToDocument() modifies cell(0,8) with header_id
- [ ] AC-08: DwgFileService.WriteTableIdsToDocument() modifies cell(row,8) with detail_id matched by product_no
- [ ] AC-09: FileDrawingDataService.WriteTableIdsAsync() exports DXF with modified TABLE
- [ ] AC-10: FileDrawingDataService.RecommendedWriteStrategy returns ExportDxf
- [ ] AC-11: Sidecar JSON is still saved as backup alongside DXF
- [ ] AC-12: BOQ page banner text updated to mention DXF export
- [ ] AC-13: Original DWG file is never modified

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-011-001 | DxfWriter SHALL support writing TableEntity | Must |
| FR-011-002 | DXF TABLE output SHALL include both subclasses | Must |
| FR-011-003 | DXF TABLE output SHALL preserve dimensions | Must |
| FR-011-004 | DXF TABLE output SHALL preserve cell content | Must |
| FR-011-005 | DXF TABLE output SHALL preserve table style | Must |
| FR-011-006 | DXF output SHALL be readable by AutoCAD 2010+ | Must |
| FR-011-007 | Cell values SHALL be modifiable in memory | Must |
| FR-011-013 | DwgFileService SHALL provide WriteTableIdsToDocument() | Must |
| FR-011-014 | FileDrawingDataService SHALL use DXF export for writeback | Must |
| FR-011-015 | RecommendedWriteStrategy SHALL return ExportDxf | Must |
| FR-011-016 | BOQ page banner SHALL indicate DXF export | Must |
| FR-011-017 | Sidecar JSON SHALL remain as backup | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-011-01-01 | Implement writeTableEntity() in DxfSectionWriterBase.Entities.cs | `ACadSharp/IO/DXF/DxfStreamWriter/DxfSectionWriterBase.Entities.cs` | L |
| TASK-011-01-02 | Remove TABLE from isEntitySupported block, add case before Insert | `ACadSharp/IO/DXF/DxfStreamWriter/DxfSectionWriterBase.Entities.cs` | S |
| TASK-011-01-03 | Add WriteTableIdsToDocument() to IDwgFileService interface | `OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs` | S |
| TASK-011-01-04 | Implement WriteTableIdsToDocument() in DwgFileService | `OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` | M |
| TASK-011-01-05 | Update FileDrawingDataService.WriteTableIdsAsync() for DXF export | `OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` | M |
| TASK-011-01-06 | Update RecommendedWriteStrategy and BOQPage banner text | `FileDrawingDataService.cs`, `BOQPage.xaml` | S |
| TASK-011-01-07 | Tests: writeTableEntity DXF roundtrip, cell write, integration | `tests/` | M |
| TASK-011-01-08 | Build and run all tests | Solution | M |

## Dependencies
- Depends on: US-010-03 (file-based writeback infrastructure, SidecarIdStore)
- Depends on: Sprint 10 complete (IDrawingDataService, FileDrawingDataService, DwgFileService)
- Blocks: US-011-02 (DWG writer builds on DXF writer patterns)

## Notes
- The DXF TABLE writer mirrors the reader at `DxfSectionReaderBase.cs:408-646` but writes instead of reads.
- `TableEntity` extends `Insert`, so the `case TableEntity` MUST appear BEFORE `case Insert` in the writer's switch statement — otherwise it falls through to the Insert handler and loses TABLE data.
- The `writeInsert()` method's block reference + attribute logic can be partially reused for the TABLE's Insert base class data.
- Cell content structure: each cell has `Contents` list; `Contents[0].Value.Text` holds the string value.
- DXF group codes follow a specific sequence: 301 "CELL_VALUE" → 90 (type) → 1 (text) → 304 "ACVALUE_END".
- This is Phase 1 — DXF only. Phase 2 (US-011-02) adds native DWG writing.
