# TASKS: US-011-01 — DXF TABLE Entity Writer

> **Parent US**: [US-011-01](US-011-01-dxf-table-writer.md)
> **Parent FR**: [FR-011](FR-011-table-entity-writer.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 3S + 4M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] Sprint 10 complete (IDrawingDataService, FileDrawingDataService, DwgFileService, SidecarIdStore)
- [ ] ACadSharp submodule functional at `csharp/libs/ACadSharp/`
- [ ] DxfWriter works for non-TABLE entities (confirmed)

## Acceptance Criteria
- [ ] AC-01: DxfWriter outputs TableEntity
- [ ] AC-02: DXF TABLE includes AcDbBlockReference + AcDbTable subclasses
- [ ] AC-03: DXF TABLE preserves row/column dimensions
- [ ] AC-04: DXF TABLE preserves cell text content
- [ ] AC-05: DXF TABLE preserves table style handle
- [ ] AC-06: DXF output readable by AutoCAD 2010+
- [ ] AC-07: WriteTableIdsToDocument() writes header_id to cell(0,8)
- [ ] AC-08: WriteTableIdsToDocument() writes detail_id matched by product_no
- [ ] AC-09: FileDrawingDataService exports DXF with modified TABLE
- [ ] AC-10: RecommendedWriteStrategy returns ExportDxf
- [ ] AC-11: Sidecar JSON still saved as backup
- [ ] AC-12: BOQ page banner mentions DXF export
- [ ] AC-13: Original DWG never modified

---

## TASK-011-01-01: Implement writeTableEntity() in DxfSectionWriterBase.Entities.cs

| Field | Value |
|-------|-------|
| Target | `csharp/libs/ACadSharp/src/ACadSharp/IO/DXF/DxfStreamWriter/DxfSectionWriterBase.Entities.cs` |
| Estimate | L |
| Depends On | — |
| Blocks | TASK-011-01-02 |

### What to do
- Implement `writeTableEntity(TableEntity table)` method (~120 lines)
- Write sequence mirrors the DXF reader at `DxfSectionReaderBase.cs:408-646`:
  1. **Insert base data** — reuse pattern from `writeInsert()`:
     - `DxfCode.Subclass` → `AcDbBlockReference`
     - Block name (code 2), insert point (code 10), scale (41/42/43), rotation (50)
     - Normal vector (210)
  2. **AcDbTable subclass**:
     - `DxfCode.Subclass` → `AcDbTable`
     - Code 342: Table style handle (`table.Style`)
     - Code 91: Row count (`table.Rows.Count`)
     - Code 92: Column count (`table.Columns.Count`)
  3. **Row heights**: For each row → code 141 (height)
  4. **Column widths**: For each column → code 142 (width)
  5. **Cell data**: For each row → each column:
     - Code 171: cell type (1=text, 2=block)
     - Code 172: edge flags
     - Code 173: merged value
     - Code 174: auto fit
     - Code 175/176: border width/height
     - Per-cell style override codes (63, 64, 69, 65, 66, 68, 140, 279, 275, 276, 278, 283)
     - Cell content (if cell has contents):
       - Code 301: `"CELL_VALUE"`
       - Code 90: value data type
       - Code 1: text value (for string type)
       - Code 300: format string (if present)
       - Code 304: `"ACVALUE_END"`
  6. **Attributes** — if `table.HasAttributes`, write attribute entities + SEQEND (reuse from `writeInsert`)

### How to verify
- [ ] Method compiles without errors
- [ ] DXF output contains `AcDbTable` section for TABLE entities
- [ ] Row/column counts match original (AC-03)
- [ ] Cell text values preserved (AC-04)

---

## TASK-011-01-02: Remove TABLE from isEntitySupported, add case before Insert

| Field | Value |
|-------|-------|
| Target | `csharp/libs/ACadSharp/src/ACadSharp/IO/DXF/DxfStreamWriter/DxfSectionWriterBase.Entities.cs` |
| Estimate | S |
| Depends On | TASK-011-01-01 |
| Blocks | TASK-011-01-07 |

### What to do
- In `isEntitySupported()` (line ~143), remove `case TableEntity:` from the block list
- In `writeEntity<T>()` switch statement (line ~45), add BEFORE `case Insert`:
  ```csharp
  case TableEntity table:
      this.writeTableEntity(table);
      break;
  ```
- **Critical**: `TableEntity` extends `Insert`. If the case appears AFTER `case Insert`, the polymorphic match will hit Insert first and TABLE data will be lost.

### How to verify
- [ ] `isEntitySupported(new TableEntity())` returns true (AC-01)
- [ ] TABLE entities routed to `writeTableEntity()` not `writeInsert()`
- [ ] Existing Insert entities still work correctly

---

## TASK-011-01-03: Add WriteTableIdsToDocument() to IDwgFileService interface

| Field | Value |
|-------|-------|
| Target | `csharp/OdooAutoCADIntegration/src/OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs` |
| Estimate | S |
| Depends On | — |
| Blocks | TASK-011-01-04 |

### What to do
- Add interface method:
  ```csharp
  /// <summary>
  /// Modifies TABLE cell values in the loaded document for a specific layout.
  /// Writes header_id to cell(0,8) and detail_id to cell(row,8) matched by product_no in col 1.
  /// </summary>
  void WriteTableIdsToDocument(string layoutName, string headerId, IList<WritebackDetail> details);
  ```
- `WritebackDetail` should already exist from Sprint 10 (in `IDrawingDataService.cs`)
- If not, verify the type and import it

### How to verify
- [ ] Interface compiles
- [ ] Method signature matches implementation expectations

---

## TASK-011-01-04: Implement WriteTableIdsToDocument() in DwgFileService

| Field | Value |
|-------|-------|
| Target | `csharp/OdooAutoCADIntegration/src/OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` |
| Estimate | M |
| Depends On | TASK-011-01-03 |
| Blocks | TASK-011-01-05 |

### What to do
- Implement the method:
  ```csharp
  public void WriteTableIdsToDocument(string layoutName, string headerId, IList<WritebackDetail> details)
  {
      if (_document == null) throw new InvalidOperationException("No document loaded");

      // Find layout
      var layout = _document.Layouts.FirstOrDefault(l => l.Name == layoutName);
      if (layout == null) return;

      // Find 9-column TABLE entities in layout
      var tables = layout.Entities.OfType<TableEntity>()
          .Where(t => t.Columns.Count == 9);

      foreach (var table in tables)
      {
          // Write header_id to cell(0, 8)
          SetCellText(table, 0, 8, headerId);

          // Write detail_ids matched by product_no in col 1
          for (int row = 2; row < table.Rows.Count; row++) // data rows start at row 2
          {
              var productNo = GetCellText(table, row, 1);
              if (string.IsNullOrEmpty(productNo)) continue;

              var detail = details.FirstOrDefault(d => d.ProductNo == productNo);
              if (detail != null)
              {
                  SetCellText(table, row, 8, detail.DetailId);
              }
          }
      }
  }

  private void SetCellText(TableEntity table, int row, int col, string value)
  {
      var cell = table.Cells[row, col];
      if (cell.Contents.Count > 0)
      {
          cell.Contents[0].Value.Text = value;
      }
      else
      {
          // Create new content entry
          var content = new CellContent();
          content.Value = new CellValue { Text = value };
          cell.Contents.Add(content);
      }
  }

  private string? GetCellText(TableEntity table, int row, int col)
  {
      var cell = table.Cells[row, col];
      if (cell.Contents.Count > 0)
          return cell.Contents[0].Value?.Text;
      return null;
  }
  ```
- Data rows start at row 2 (row 0 = title, row 1 = header) — matches Python `if i > 1`
- Product matching uses column 1 (product_no)

### How to verify
- [ ] Cell(0,8) contains header_id after call (AC-07)
- [ ] Cell(row,8) contains detail_id for matching product_no (AC-08)
- [ ] Non-matching rows are not modified
- [ ] Throws if no document loaded

---

## TASK-011-01-05: Update FileDrawingDataService.WriteTableIdsAsync() for DXF export

| Field | Value |
|-------|-------|
| Target | `csharp/OdooAutoCADIntegration/src/OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` |
| Estimate | M |
| Depends On | TASK-011-01-04 |
| Blocks | TASK-011-01-06 |

### What to do
- Modify `WriteTableIdsAsync()` to:
  1. Write cell values in memory via `_dwgFileService.WriteTableIdsToDocument()`
  2. Export modified document as DXF via `_dwgFileService.SaveAsDxfAsync()`
  3. Keep sidecar JSON as backup (existing behavior)
  4. Return `WritebackResult` with DXF path and `WriteStrategy.ExportDxf`
  ```csharp
  public async Task<WritebackResult> WriteTableIdsAsync(
      string layoutName, string headerId, IList<WritebackDetail> details)
  {
      // 1. Modify cells in memory
      _dwgFileService.WriteTableIdsToDocument(layoutName, headerId, details);

      // 2. Export as DXF
      var dxfPath = Path.ChangeExtension(_dwgFileService.LoadedFilePath!, ".dxf");
      var dxfSuccess = await _dwgFileService.SaveAsDxfAsync(dxfPath);

      // 3. Backup to sidecar JSON (existing logic)
      var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath!)
                    ?? new SidecarData();
      sidecar.SourceDwg = Path.GetFileName(_dwgFileService.LoadedFilePath!);
      sidecar.ModifiedAt = DateTime.UtcNow;
      var sidecarLayout = sidecar.Layouts.FirstOrDefault(l => l.Name == layoutName)
                          ?? new SidecarLayout { Name = layoutName };
      if (!sidecar.Layouts.Contains(sidecarLayout))
          sidecar.Layouts.Add(sidecarLayout);
      sidecarLayout.HeaderId = headerId;
      sidecarLayout.Details = details.Select(d => new SidecarDetail
      {
          ProductNo = d.ProductNo, DetailId = d.DetailId
      }).ToList();
      await _sidecarStore.SaveAsync(_dwgFileService.LoadedFilePath!, sidecar);

      if (dxfSuccess)
          return new WritebackResult(true, $"ID 已寫入 DXF: {dxfPath}", WriteStrategy.ExportDxf, dxfPath);
      else
          return new WritebackResult(true, "ID 已儲存至 sidecar JSON (DXF 匯出失敗)", WriteStrategy.SidecarJson);
  }
  ```

### How to verify
- [ ] DXF file created alongside DWG (AC-09)
- [ ] Sidecar JSON also created (AC-11)
- [ ] Original DWG not modified (AC-13)
- [ ] Returns ExportDxf strategy on success

---

## TASK-011-01-06: Update RecommendedWriteStrategy and BOQPage banner text

| Field | Value |
|-------|-------|
| Target | `FileDrawingDataService.cs`, `Views/Pages/BOQPage.xaml` |
| Estimate | S |
| Depends On | TASK-011-01-05 |
| Blocks | — |

### What to do
- In `FileDrawingDataService`:
  ```csharp
  public WriteStrategy RecommendedWriteStrategy => WriteStrategy.ExportDxf;
  ```
  (was `SidecarJson`)
- In `BOQPage.xaml`, update the yellow file mode banner text:
  ```
  Before: "檔案模式：ID 回寫將儲存至 sidecar JSON 檔案"
  After:  "檔案模式：ID 將寫入 DXF 匯出檔案（同時備份至 sidecar JSON）"
  ```

### How to verify
- [ ] RecommendedWriteStrategy returns ExportDxf (AC-10)
- [ ] BOQ page banner shows updated DXF message (AC-12)

---

## TASK-011-01-07: Tests — writeTableEntity DXF roundtrip, cell write, integration

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/` |
| Estimate | M |
| Depends On | TASK-011-01-02, TASK-011-01-04, TASK-011-01-05 |
| Blocks | TASK-011-01-08 |

### What to do
- Add ~6 new test methods:
  ```csharp
  // DwgFileService cell modification
  [Fact] WriteTableIdsToDocument_SetsHeaderIdInCell08()
  [Fact] WriteTableIdsToDocument_SetsDetailIdByProductNo()
  [Fact] WriteTableIdsToDocument_SkipsNonMatchingRows()

  // FileDrawingDataService integration
  [Fact] WriteTableIdsAsync_ExportsDxfFile()
  [Fact] WriteTableIdsAsync_CreatesSidecarJsonBackup()
  [Fact] RecommendedWriteStrategy_ReturnsExportDxf()
  ```
- For DXF roundtrip tests, create a minimal TABLE in memory:
  ```csharp
  var table = new TableEntity();
  table.InsertRows(0, 10, 3);   // 3 rows
  table.InsertColumns(0, 20, 9); // 9 columns
  // Set cell values, write to DXF, read back, verify
  ```
- Mock `IDwgFileService` for FileDrawingDataService unit tests
- Use temp directory for DXF output files

### How to verify
- [ ] All 6 tests pass
- [ ] Cell values verified after write (AC-07, AC-08)
- [ ] DXF file created (AC-09)
- [ ] Sidecar JSON created (AC-11)

---

## TASK-011-01-08: Build and run all tests

| Field | Value |
|-------|-------|
| Target | Solution |
| Estimate | M |
| Depends On | TASK-011-01-07 |
| Blocks | — |

### What to do
- Build entire solution: `dotnet build OdooAutoCADIntegration.sln`
- Run all tests: `dotnet test` — verify 270 existing + ~6 new = ~276 pass
- Verify no new warnings in the TABLE writer code
- Verify existing DXF export functionality (non-TABLE entities) is unaffected

### How to verify
- [ ] Build succeeds with 0 errors
- [ ] All ~276 tests pass
- [ ] No new warnings in ACadSharp writer code
- [ ] Existing tests unchanged

---

## Dependency Graph
```
TASK-011-01-01 (writeTableEntity impl)
    |
    +---> TASK-011-01-02 (enable in isEntitySupported + switch)
              |
              +---> TASK-011-01-07 (tests)

TASK-011-01-03 (IDwgFileService interface)
    |
    +---> TASK-011-01-04 (DwgFileService impl)
              |
              +---> TASK-011-01-05 (FileDrawingDataService DXF export)
                        |
                        +---> TASK-011-01-06 (strategy + banner)
                        |
                        +---> TASK-011-01-07 (tests)
                                  |
                                  +---> TASK-011-01-08 (build & test all)
```
