# TASKS: US-010-03 — File-Based Writeback

> **Parent US**: [US-010-03](US-010-03-file-based-writeback.md)
> **Parent FR**: [FR-010](FR-010-dual-mode-autocad.md)
> **Priority**: P1
> **Tasks**: 7 | **Effort**: 1S + 6M
> **Status**: Not Started

## Prerequisites
- [ ] US-010-02 complete (file read infrastructure, DwgFileService, FileDrawingDataService)
- [ ] ACadSharp submodule functional with DxfWriter

## Acceptance Criteria
- [ ] AC-01: Block attribute writes produce a DXF file via DxfWriter
- [ ] AC-02: Table ID writeback creates sidecar JSON
- [ ] AC-03: Sidecar JSON contains version, source_dwg, modified_at, layouts
- [ ] AC-04: Original DWG is never modified
- [ ] AC-05: SupportsWrite returns true (for attributes via DXF)
- [ ] AC-06: RecommendedWriteStrategy returns SidecarJson for table IDs
- [ ] AC-07: BOQ page shows info banner in file mode
- [ ] AC-08: Sidecar JSON loadable when switching to COM mode
- [ ] AC-09: SidecarIdStore handles concurrent access safely

---

## TASK-010-03-01: Implement SidecarIdStore — JSON sidecar read/write

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/SidecarIdStore.cs` |
| Estimate | M |
| Depends On | — |
| Blocks | TASK-010-03-02 |

### What to do
- Create `SidecarIdStore` class:
  ```csharp
  public class SidecarIdStore
  {
      public Task<SidecarData?> LoadAsync(string dwgFilePath);
      public Task SaveAsync(string dwgFilePath, SidecarData data);
      public static string GetSidecarPath(string dwgFilePath);
      public Task<bool> ExistsAsync(string dwgFilePath);
  }
  ```
- Define `SidecarData` model:
  ```csharp
  public class SidecarData
  {
      public int Version { get; set; } = 1;
      public string SourceDwg { get; set; } = string.Empty;
      public DateTime ModifiedAt { get; set; }
      public List<SidecarLayout> Layouts { get; set; } = new();
      public Dictionary<string, string> Attributes { get; set; } = new();
  }
  public class SidecarLayout
  {
      public string Name { get; set; } = string.Empty;
      public string HeaderId { get; set; } = string.Empty;
      public List<SidecarDetail> Details { get; set; } = new();
  }
  public class SidecarDetail
  {
      public string ProductNo { get; set; } = string.Empty;
      public string DetailId { get; set; } = string.Empty;
  }
  ```
- `GetSidecarPath`: `Path.ChangeExtension(dwgFilePath, null) + ".boq-ids.json"`
- Use `System.Text.Json` with `JsonSerializerOptions { WriteIndented = true }`
- `SaveAsync`: lock via `SemaphoreSlim(1,1)` for concurrent access safety
- `LoadAsync`: return null if file doesn't exist

### How to verify
- [ ] Sidecar JSON created alongside DWG file (AC-02)
- [ ] JSON contains all required fields (AC-03)
- [ ] Concurrent writes don't corrupt file (AC-09)
- [ ] Load returns null for non-existent sidecar

---

## TASK-010-03-02: Implement FileDrawingDataService — write operations

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` |
| Estimate | M |
| Depends On | TASK-010-03-01, TASK-010-02-05 |
| Blocks | TASK-010-03-03 |

### What to do
- Implement write methods in `FileDrawingDataService`:
  ```csharp
  public bool SupportsWrite => true; // DXF export works
  public WriteStrategy RecommendedWriteStrategy => WriteStrategy.SidecarJson;

  public async Task<WritebackResult> WriteTableIdsAsync(
      string layoutName, string headerId, IList<WritebackDetail> details)
  {
      // Primary: sidecar JSON
      var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath!)
                    ?? new SidecarData();
      sidecar.SourceDwg = Path.GetFileName(_dwgFileService.LoadedFilePath!);
      sidecar.ModifiedAt = DateTime.UtcNow;
      // Update/add layout entry
      var layout = sidecar.Layouts.FirstOrDefault(l => l.Name == layoutName)
                   ?? new SidecarLayout { Name = layoutName };
      layout.HeaderId = headerId;
      layout.Details = details.Select(d => new SidecarDetail
      {
          ProductNo = d.ProductNo, DetailId = d.DetailId
      }).ToList();
      // Save
      await _sidecarStore.SaveAsync(_dwgFileService.LoadedFilePath!, sidecar);
      return new WritebackResult(true, "ID 已儲存至 sidecar JSON", WriteStrategy.SidecarJson);
  }

  public async Task<List<string>> SetAttributeValuesAsync(
      Dictionary<string, string> attributes, string? layoutName)
  {
      // Store in sidecar for now; DXF export in SaveAsync
      var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath!)
                    ?? new SidecarData();
      foreach (var (key, value) in attributes)
          sidecar.Attributes[key] = value;
      sidecar.ModifiedAt = DateTime.UtcNow;
      await _sidecarStore.SaveAsync(_dwgFileService.LoadedFilePath!, sidecar);
      return attributes.Keys.ToList();
  }
  ```
- `SaveAsync()` triggers DXF export if attribute modifications exist

### How to verify
- [ ] WriteTableIdsAsync creates sidecar JSON (AC-02)
- [ ] SetAttributeValuesAsync stores attributes in sidecar (AC-01)
- [ ] SupportsWrite returns true (AC-05)
- [ ] RecommendedWriteStrategy returns SidecarJson (AC-06)

---

## TASK-010-03-03: Implement FileDrawingDataService.SaveAsDxf() for DXF export

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` |
| Estimate | M |
| Depends On | TASK-010-03-02 |
| Blocks | — |

### What to do
- Implement `SaveAsync()` in `FileDrawingDataService`:
  ```csharp
  public async Task<bool> SaveAsync()
  {
      if (!IsReady) return false;

      // Apply attribute modifications from sidecar to CadDocument
      var sidecar = await _sidecarStore.LoadAsync(_dwgFileService.LoadedFilePath!);
      if (sidecar?.Attributes?.Count > 0)
      {
          _dwgFileService.ApplyAttributesToDocument(sidecar.Attributes);
      }

      // Export as DXF (never overwrite original DWG)
      var dxfPath = Path.ChangeExtension(_dwgFileService.LoadedFilePath!, ".dxf");
      return await _dwgFileService.SaveAsDxfAsync(dxfPath);
  }
  ```
- In `DwgFileService`, implement `SaveAsDxfAsync(outputPath)`:
  1. Apply pending attribute modifications to `_document` in memory
  2. Use ACadSharp `DxfWriter` to export
  3. Write as ASCII DXF (broadest compatibility)
  4. Return success/failure
- Implement `ApplyAttributesToDocument(attributes)`:
  1. Find attribute blocks in each layout
  2. Set attribute tag values from dictionary
  3. This modifies the in-memory document only (original DWG unchanged)

### How to verify
- [ ] DXF file created alongside DWG (AC-01)
- [ ] Original DWG not modified (AC-04)
- [ ] Attributes applied to in-memory document before export
- [ ] DXF readable by AutoCAD/AutoCAD LT

---

## TASK-010-03-04: Refactor BOQViewModel for sidecar writeback fallback

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQViewModel.cs` |
| Estimate | M |
| Depends On | TASK-010-03-02, TASK-010-04-02 |
| Blocks | — |

### What to do
- In BOQViewModel's push/writeback logic, check `_drawingDataService.RecommendedWriteStrategy`:
  ```csharp
  if (_drawingDataService.RecommendedWriteStrategy == WriteStrategy.SidecarJson)
  {
      // Use sidecar writeback
      var result = await _drawingDataService.WriteTableIdsAsync(layout, headerId, details);
      LastPushResult = result.Message;
  }
  else
  {
      // Use direct COM writeback (existing logic)
      var result = await _drawingDataService.WriteTableIdsAsync(layout, headerId, details);
      // COM backend writes directly to AutoCAD
  }
  ```
- Actually both paths call the same `WriteTableIdsAsync` — the backend handles the strategy
- Add UI feedback: when sidecar strategy, show `"ID 已儲存至 sidecar JSON (非直接寫入 AutoCAD)"`
- The ViewModel logic is mode-agnostic; only status messages differ

### How to verify
- [ ] BOQViewModel calls IDrawingDataService for writeback
- [ ] Sidecar writeback shows appropriate message
- [ ] COM writeback behavior unchanged

---

## TASK-010-03-05: Add file mode info banner to BOQ page

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQPage.xaml` |
| Estimate | S |
| Depends On | — |
| Blocks | — |

### What to do
- Add info banner at the top of BOQ page, visible only in file mode:
  ```xml
  <Border Background="#FFF3CD" BorderBrush="#FFC107" BorderThickness="1"
          CornerRadius="4" Padding="12,8" Margin="0,0,0,12"
          Visibility="{Binding IsFileMode, Converter={StaticResource BoolToVisibilityConverter}}">
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="ℹ" FontSize="16" Margin="0,0,8,0"/>
      <TextBlock TextWrapping="Wrap">
        <Run Text="檔案模式："/>
        <Run Text="ID 回寫將儲存至 sidecar JSON 檔案"/>
        <Run Text=" ({drawing}.boq-ids.json)，不會修改原始 DWG。" Foreground="Gray"/>
      </TextBlock>
    </StackPanel>
  </Border>
  ```
- In BOQViewModel, add `IsFileMode` computed property:
  ```csharp
  public bool IsFileMode => _drawingDataService.Mode == AutoCADOperationMode.File;
  ```

### How to verify
- [ ] Banner visible when in file mode (AC-07)
- [ ] Banner hidden when in COM mode
- [ ] Text explains sidecar writeback clearly

---

## TASK-010-03-06: Tests — SidecarIdStore roundtrip and concurrent access

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/SidecarIdStoreTests.cs` |
| Estimate | M |
| Depends On | TASK-010-03-01 |
| Blocks | — |

### What to do
- Create `SidecarIdStoreTests.cs`:
  ```csharp
  [Fact] GetSidecarPath_ReturnsCorrectPath()
  [Fact] SaveAndLoad_Roundtrip_PreservesAllData()
  [Fact] Load_NonExistentFile_ReturnsNull()
  [Fact] Save_CreatesJsonFile()
  [Fact] Save_OverwritesExistingFile()
  [Fact] Exists_ReturnsTrue_WhenFileExists()
  [Fact] Exists_ReturnsFalse_WhenNoFile()
  [Fact] ConcurrentSaves_DoNotCorruptFile()
  ```
- Use temporary directory for test files
- Verify JSON structure matches expected schema
- Concurrent test: multiple Task.Run saves, verify final file is valid JSON

### How to verify
- [ ] All 8 tests pass
- [ ] Roundtrip preserves all fields (AC-03)
- [ ] Concurrent access safe (AC-09)

---

## TASK-010-03-07: Tests — FileDrawingDataService write operations

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/FileDrawingDataServiceTests.cs` |
| Estimate | M |
| Depends On | TASK-010-03-02, TASK-010-03-03 |
| Blocks | — |

### What to do
- Create `FileDrawingDataServiceTests.cs`:
  ```csharp
  [Fact] WriteTableIdsAsync_CreatesSidecarJson()
  [Fact] WriteTableIdsAsync_UpdatesExistingSidecar()
  [Fact] SetAttributeValuesAsync_StoresInSidecar()
  [Fact] SaveAsync_ExportsDxfFile()
  [Fact] SaveAsync_DoesNotModifyOriginalDwg()
  [Fact] SupportsWrite_ReturnsTrue()
  [Fact] RecommendedWriteStrategy_ReturnsSidecarJson()
  ```
- Mock `IDwgFileService` and `SidecarIdStore` for unit tests
- Use temporary files for integration-level DXF export tests

### How to verify
- [ ] All 7 tests pass
- [ ] Write operations create sidecar JSON (AC-02)
- [ ] DXF export creates file (AC-01)
- [ ] Original DWG unchanged (AC-04)

---

## Dependency Graph
```
TASK-010-03-01 (SidecarIdStore)
    |
    +---> TASK-010-03-02 (FileDrawingDataService write ops)
    |         |
    |         +---> TASK-010-03-03 (DXF export)
    |         |
    |         +---> TASK-010-03-04 (BOQViewModel sidecar fallback)
    |
    +---> TASK-010-03-06 (SidecarIdStore tests)

TASK-010-03-05 (BOQ page info banner — independent)

TASK-010-03-07 (FileDrawingDataService write tests)
```
