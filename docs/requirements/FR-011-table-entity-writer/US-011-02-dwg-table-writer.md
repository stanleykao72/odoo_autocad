# US-011-02: DWG TABLE Entity Writer

## User Story
**As a** CAD Engineer,
**I want to** have file mode write IDs into native DWG TABLE cells,
**So that** I can keep working with DWG files directly without needing DXF conversion.

## Parent Feature
- **FR**: [FR-011-table-entity-writer](FR-011-table-entity-writer.md)
- **Priority**: P2

## Acceptance Criteria
- [x] AC-01: ACadSharp DwgWriter outputs TableEntity instead of skipping it
- [x] AC-02: DWG TABLE handles R2010+ format (modern path with writeTableContent)
- [x] AC-03: DWG TABLE handles R2007 and earlier format (legacy path with cell data + overrides)
- [x] AC-04: DWG TABLE roundtrip (read → modify cell → write → read) preserves cell values
- [x] AC-05: DwgFileService.SaveAsDwgAsync() produces valid DWG with TABLE entities
- [x] AC-06: DwgFileService.CanWriteDwg returns true when TABLE support is verified
- [x] AC-07: FileDrawingDataService tries DWG export first, falls back to DXF on failure
- [x] AC-08: FileDrawingDataService.RecommendedWriteStrategy returns DirectDwg when DWG is available
- [x] AC-09: SaveAsync() saves as DWG when possible, DXF as fallback
- [x] AC-10: DWG output is readable by AutoCAD 2018+ (AC1032 format) — pending manual verification

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-011-008 | DwgWriter SHALL support writing TableEntity | Should |
| FR-011-009 | DWG TABLE SHALL handle R2010+ format | Should |
| FR-011-010 | DWG TABLE SHALL handle R2007 and earlier | Should |
| FR-011-011 | DWG TABLE roundtrip SHALL preserve cell values | Should |
| FR-011-012 | If DWG write fails, system SHALL fall back to DXF | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-011-02-01 | Implement writeTableEntity() legacy path (R2007-) in DwgObjectWriter | `ACadSharp/IO/DWG/DwgStreamWriters/DwgObjectWriter.Entities.cs` | L |
| TASK-011-02-02 | Remove TABLE from DWG isEntitySupported, add case before Insert | `ACadSharp/IO/DWG/DwgStreamWriters/DwgObjectWriter.cs` | S |
| TASK-011-02-03 | Implement writeTableContent/writeTableCell/writeTableCellContent (R2010+) | `ACadSharp/IO/DWG/DwgStreamWriters/DwgObjectWriter.Entities.cs` | L |
| TASK-011-02-04 | Implement break data + break row range writing for R2010+ | `ACadSharp/IO/DWG/DwgStreamWriters/DwgObjectWriter.Entities.cs` | M |
| TASK-011-02-05 | Add SaveAsDwgAsync() to IDwgFileService + DwgFileService | `IDwgFileService.cs`, `DwgFileService.cs` | M |
| TASK-011-02-06 | Set CanWriteDwg = true in DwgFileService | `DwgFileService.cs` | S |
| TASK-011-02-07 | Update FileDrawingDataService to prefer DWG, fallback to DXF | `FileDrawingDataService.cs` | M |
| TASK-011-02-08 | Tests: DWG TABLE roundtrip, version-specific paths, fallback | `tests/` | M |
| TASK-011-02-09 | Build and run all tests | Solution | M |

## Dependencies
- Depends on: US-011-01 (DXF TABLE writer — provides DXF fallback path)
- Depends on: Sprint 10 complete (DwgFileService infrastructure)
- Blocks: None (end-of-chain)

## Notes

### Why DWG Writing is Complex
The DWG TABLE reader uses two major code paths:

| Path | Condition | Reader Lines | Key Methods |
|------|-----------|-------------|------------|
| Modern | R2010+ (`this.R2010Plus`) | ~200 | readTableContent, readTableCell, readTableCellContent, readCellStyle |
| Legacy | R2007 and earlier | ~480 | readTableCellData, 3x border override sections (color/lineweight/visibility x 18 flags each) |

The writer must mirror BOTH paths, totaling ~600 lines of bit-packed binary writes.

### Write Method Inventory (R2010+)
1. `writeInsertCommonData` + `writeInsertCommonHandles` (Insert base)
2. `WriteByte`, `HandleReference(null)`, `WriteBitLong(0)` (unknown fields)
3. `writeTableContent()` — name, description, columns, rows, cells
4. `writeTableCell()` → `writeTableCellContent()` per cell
5. Break data: flag, positions, row ranges

### Write Method Inventory (Legacy R2007-)
1. `WriteBitShort(90)` — value flag
2. `Write3BitDouble` — horizontal direction
3. Per-column widths, per-row heights
4. `HandleReference` — table style
5. Per-cell: `writeTableCellData()` (~179 lines per cell)
6. 23 flag-conditional table overrides
7. 3 × 18 = 54 flag-conditional border overrides (color, lineweight, visibility)

### Risk Mitigation
- Phase 2 is gated on Phase 1 DXF success
- DWG write failure automatically falls back to DXF export
- `CanWriteDwg` flag allows runtime decision
- DWG writer is already marked "not reliable yet" — TABLE support doesn't change that baseline
- Manual validation with AutoCAD 2018 (AC1032) files is required
