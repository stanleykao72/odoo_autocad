# FR-011: TABLE Entity Writer — Native ID Writeback in File Mode

> **Document Version**: 1.0
> **Last Updated**: 2026-02-13
> **Status**: Complete (Phase 1 + Phase 2)
> **Priority**: P1

## 1. Overview

File mode (FR-010) can read TABLE entities from DWG files via ACadSharp but cannot write them back. Both `DxfWriter` and `DwgWriter` explicitly skip `TableEntity` in their `isEntitySupported()` methods. This means after pushing BOQ to Odoo, the returned `header_id` and `detail_id` values cannot be written back into the drawing's TABLE cells — they only go to a sidecar JSON file.

This feature implements TABLE entity writing in ACadSharp so that File mode can:
1. Modify TABLE cell values in memory (header_id at cell(0,8), detail_id at cell(row,8))
2. Export a valid DXF file with TABLE entities preserved (Phase 1)
3. Export a valid DWG file with TABLE entities natively (Phase 2)

**Key Design Decisions**:
- ACadSharp is a Git submodule — we can modify its writers directly
- Phase 1 (DXF) is low-risk (~120 lines) because the DXF format is text-based
- Phase 2 (DWG) is high-risk (~600 lines) because the DWG format is bit-packed binary
- DXF export is the safe, validated path; DWG export is a stretch goal
- Sidecar JSON remains as a backup strategy even after TABLE writing works

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-011-01 | CAD Engineer | Have file mode write IDs directly into DXF TABLE cells | I can open the exported DXF in AutoCAD and see header_id/detail_id values in the table |
| US-011-02 | CAD Engineer | Have file mode write IDs into native DWG TABLE cells | I don't need to convert between DXF and DWG formats to see writeback results |

## 3. ACadSharp Reference

### Source Files (READ path — mirrors for WRITE)
- `ACadSharp/IO/DXF/DxfSectionReaderBase.cs` (lines 408-646) — DXF TABLE reader (~236 lines)
- `ACadSharp/IO/DWG/DwgStreamReaders/DwgObjectReader.Entities.cs` (lines 15-591) — DWG TABLE reader (~1,100 lines)
- `ACadSharp/Entities/TableEntity.cs` — TableEntity model (extends Insert)
- `ACadSharp/Entities/TableEntity.Cell.cs` — Cell model with Contents, CellType, Value

### Key Patterns
- `TableEntity` extends `Insert` — both writers must handle it BEFORE the `case Insert` in their switch statements
- Cell values accessed via `cell.Contents[0].Value.Text` (setters exist, verified)
- DXF TABLE format uses group codes: 171 (cell type), 301 "CELL_VALUE", 90 (value type), 1 (text), 304 "ACVALUE_END"
- DWG TABLE has two code paths: R2010+ (modern, ~200 lines) and Legacy R2007- (~480 lines)
- 9-column table layout: cols 0-6 are data, col 7 is HEADER_ID, col 8 is detail_id

### ACadSharp Writer Entry Points
- DXF: `DxfSectionWriterBase.Entities.cs` — `writeEntity<T>()` switch + `isEntitySupported()`
- DWG: `DwgObjectWriter.cs` (line 259) — `isEntitySupported()` + `DwgObjectWriter.Entities.cs` (line 59) switch

## 4. Functional Requirements

### DXF TABLE Writing (Phase 1)

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-011-001 | DxfWriter SHALL support writing TableEntity to DXF output | Must |
| FR-011-002 | DXF TABLE output SHALL include AcDbBlockReference + AcDbTable subclasses | Must |
| FR-011-003 | DXF TABLE output SHALL preserve row/column dimensions | Must |
| FR-011-004 | DXF TABLE output SHALL preserve cell content text values | Must |
| FR-011-005 | DXF TABLE output SHALL preserve table style reference | Must |
| FR-011-006 | DXF TABLE output SHALL be readable by AutoCAD 2010+ | Must |
| FR-011-007 | Cell values SHALL be modifiable in memory before export | Must |

### DWG TABLE Writing (Phase 2)

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-011-008 | DwgWriter SHALL support writing TableEntity to DWG output | Should |
| FR-011-009 | DWG TABLE output SHALL handle R2010+ format (modern path) | Should |
| FR-011-010 | DWG TABLE output SHALL handle R2007 and earlier format (legacy path) | Should |
| FR-011-011 | DWG TABLE roundtrip (read → modify → write → read) SHALL preserve cell values | Should |
| FR-011-012 | If DWG write fails, system SHALL fall back to DXF export | Must |

### Integration

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-011-013 | DwgFileService SHALL provide WriteTableIdsToDocument() for cell modification | Must |
| FR-011-014 | FileDrawingDataService SHALL use DXF export for TABLE writeback | Must |
| FR-011-015 | FileDrawingDataService.RecommendedWriteStrategy SHALL return ExportDxf | Must |
| FR-011-016 | BOQ page file mode banner SHALL indicate DXF export strategy | Must |
| FR-011-017 | Sidecar JSON SHALL remain as backup alongside DXF export | Should |

## 5. Data Model

### Cell Value Modification (In-Memory)
```csharp
// Modify TABLE cell value via ACadSharp API
TableEntity table = ...; // from layout entities
var cell = table.Cells[row, col];
if (cell.Contents.Count > 0)
{
    cell.Contents[0].Value.Text = "new_value";
}
else
{
    // Create new cell content
    var content = new CellContent { Value = new CellValue { Text = "new_value" } };
    cell.Contents.Add(content);
}
```

### DXF TABLE Group Code Sequence
```
100  AcDbBlockReference    (Insert base class)
  2  *T1                   (block name)
 10  x, 20 y, 30 z        (insertion point)
 41  xscale, 42 yscale, 43 zscale
 50  rotation
100  AcDbTable
342  table_style_handle
 91  row_count
 92  column_count
141  row_height (per row)
142  col_width (per column)
--- per cell ---
171  cell_type (1=text, 2=block)
301  CELL_VALUE
 90  value_data_type
  1  text_value
304  ACVALUE_END
```

## 6. API/Service Dependencies

| Service | Interface | Methods Used |
|---------|-----------|-------------|
| DWG File Service | `IDwgFileService` | `WriteTableIdsToDocument()` — new method |
| File Drawing Data | `FileDrawingDataService` | `WriteTableIdsAsync()` — enhanced to export DXF |
| ACadSharp DXF Writer | `DxfWriter` | Internal — `writeTableEntity()` added |
| ACadSharp DWG Writer | `DwgWriter` | Internal — `writeTableEntity()` added (Phase 2) |
| Sidecar Store | `SidecarIdStore` | Backup writeback (unchanged) |

## 7. Validation Rules

| Rule | Description |
|------|-------------|
| VR-011-001 | TABLE must have exactly 9 columns for ID writeback |
| VR-011-002 | Header ID written to cell(0, 8) — first data row |
| VR-011-003 | Detail ID written to cell(row, 8) where row matches product_no in col 1 |
| VR-011-004 | Empty cells should be created with text content if not present |
| VR-011-005 | DXF export produces file alongside original DWG (never overwrites) |

## 8. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| TABLE entity not found in layout | "佈局 '{name}' 中未找到 9 欄表格" | Skip layout, return error in result |
| DXF export fails | "DXF 匯出失敗: {details}" | Fall back to sidecar JSON only |
| DWG export fails (Phase 2) | "DWG 匯出失敗，改用 DXF 匯出" | Automatic fallback to DXF |
| Cell content creation fails | "無法寫入儲存格 ({row},{col}): {details}" | Log error, continue with other cells |
| File write permission denied | "無法寫入檔案: {path}" | Show error with path |

## 9. Test Strategy

### Phase 1 Tests (~6 new)
- Cell modification in memory (set text, verify read-back)
- DXF TABLE roundtrip (read DWG → write DXF → read DXF → verify TABLE preserved)
- WriteTableIdsToDocument (find 9-col table, write IDs, verify cells)
- FileDrawingDataService DXF export integration
- RecommendedWriteStrategy returns ExportDxf
- BOQ page banner text update

### Phase 2 Tests (~10 new)
- DWG TABLE roundtrip (read DWG → modify → write DWG → read → verify)
- R2010+ specific path coverage
- Legacy R2007 path coverage
- DWG→DXF fallback when DWG write fails
- Version-specific file handling

## 10. Implementation Notes

### Modified Files (ACadSharp Library)

**Phase 1 — DXF Writer**:
- `ACadSharp/IO/DXF/DxfStreamWriter/DxfSectionWriterBase.Entities.cs`
  - Remove `TableEntity` from `isEntitySupported()` block list
  - Add `case TableEntity table:` BEFORE `case Insert` in switch
  - Implement `writeTableEntity()` (~120 lines)

**Phase 2 — DWG Writer**:
- `ACadSharp/IO/DWG/DwgStreamWriters/DwgObjectWriter.cs`
  - Remove `TableEntity` from `isEntitySupported()` block list
- `ACadSharp/IO/DWG/DwgStreamWriters/DwgObjectWriter.Entities.cs`
  - Add `case TableEntity table:` BEFORE `case Insert` in switch
  - Implement `writeTableEntity()` (~600 lines, two paths: R2010+ and Legacy)

### Modified Files (Application)

- `OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs` — Add `WriteTableIdsToDocument()`
- `OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` — Implement cell modification
- `OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` — DXF export writeback
- `Views/Pages/BOQPage.xaml` — Update file mode banner text

### Architecture
```
BOQViewModel.PushToOdooCommand
    → IDrawingDataService.WriteTableIdsAsync(layout, headerId, details)
        → FileDrawingDataService.WriteTableIdsAsync()
            1. DwgFileService.WriteTableIdsToDocument()  ← modify cells in memory
            2. DwgFileService.SaveAsDxfAsync()           ← export DXF with TABLE
            3. SidecarIdStore.SaveAsync()                ← backup sidecar JSON
            → return WritebackResult(true, path, WriteStrategy.ExportDxf)
```

## 11. Dependencies

- **Depends on**: FR-010 (Dual-Mode AutoCAD) — file mode infrastructure
- **Depends on**: Sprint 10 complete — IDrawingDataService, FileDrawingDataService, DwgFileService
- **Blocks**: None (enhancement to existing functionality)

## 12. Timeline & Effort

| Phase | Tasks | Effort | Risk |
|-------|-------|--------|------|
| Phase 1: DXF TABLE Writer | 8 | 3S + 4M + 1L | Low |
| Phase 2: DWG TABLE Writer | 9 | 2S + 5M + 2L | High |
| **Total** | **17** | **5S + 9M + 3L** | Mixed |
