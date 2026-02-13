# TASKS: US-010-02 — File-Based Data Extraction

> **Parent US**: [US-010-02](US-010-02-file-based-extraction.md)
> **Parent FR**: [FR-010](FR-010-dual-mode-autocad.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 2S + 6M
> **Status**: Not Started

## Prerequisites
- [ ] US-010-01 complete (IDrawingDataService interface, mode infrastructure)
- [ ] ACadSharp submodule added (TASK-010-01-01)
- [ ] Existing IDwgReaderService + DwgReaderService available

## Acceptance Criteria
- [ ] AC-01: User can load a DWG file via file picker in AutoCAD page
- [ ] AC-02: Layouts are extracted from DWG file excluding "Model"
- [ ] AC-03: Table data (9-column format) is extracted from layout tables
- [ ] AC-04: Block attributes are extracted from attribute blocks
- [ ] AC-05: Header IDs are read from table cell(0,8)
- [ ] AC-06: PR number is extracted from block reference text/attributes
- [ ] AC-07: Extracted data matches COM mode output for the same DWG
- [ ] AC-08: File info (name, version, layout count) is displayed
- [ ] AC-09: ACadSharp table read failures return empty data

---

## TASK-010-02-01: Define IDwgFileService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs` |
| Estimate | S |
| Depends On | TASK-010-01-02 |
| Blocks | TASK-010-02-02, TASK-010-02-05 |

### What to do
- Create `IDwgFileService` interface that extends `IDwgReaderService`:
  ```csharp
  public interface IDwgFileService : IDwgReaderService
  {
      // File info
      bool IsFileLoaded { get; }
      string? LoadedFilePath { get; }
      string? DwgVersion { get; }
      bool CanWriteDwg { get; }

      // Load / Unload
      Task<bool> LoadFileAsync(string filePath);
      Task UnloadFileAsync();

      // Table operations (new — extends read-only base)
      Task<List<TableData>> ExtractTableDataAsync(string layoutName);
      Task<IReadOnlyList<string>> GetHeaderIdsAsync();
      Task<string> GetPRNumberAsync();

      // Write operations
      Task<bool> SaveAsDxfAsync(string outputPath);
      Task<bool> WriteAttributesAsync(string layoutName, Dictionary<string, string> attributes);
  }
  ```
- `TableData` record: `LayoutName`, `HeaderId`, `Rows` (list of row dictionaries)
- Reuse existing `LayoutInfo` from IDwgReaderService for layout listing

### How to verify
- [ ] `IDwgFileService` extends `IDwgReaderService`
- [ ] All table/write methods declared
- [ ] Solution builds without errors

---

## TASK-010-02-02: Implement DwgFileService — file info, version detection, CanWriteDwg

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` |
| Estimate | M |
| Depends On | TASK-010-02-01, TASK-010-01-01 |
| Blocks | TASK-010-02-03, TASK-010-02-04 |

### What to do
- Create `DwgFileService` extending existing `DwgReaderService` and implementing `IDwgFileService`
- Implement file management:
  ```csharp
  private CadDocument? _document;
  private string? _loadedFilePath;

  public bool IsFileLoaded => _document != null;
  public string? LoadedFilePath => _loadedFilePath;
  public string? DwgVersion => _document?.Header?.Version.ToString();
  public bool CanWriteDwg => false; // DwgWriter not reliable yet
  ```
- `LoadFileAsync(filePath)`:
  1. Validate file exists and has `.dwg` extension
  2. Open with `DwgReader` and store `_document`
  3. Store file path
  4. Return true on success, false on failure
- `UnloadFileAsync()`:
  1. Dispose `_document` if IDisposable
  2. Clear `_document` and `_loadedFilePath`
- Version detection: read from `_document.Header.Version` (e.g., AC1032 = AutoCAD 2018)

### How to verify
- [ ] `LoadFileAsync` opens a DWG and sets IsFileLoaded = true (AC-01)
- [ ] `DwgVersion` returns version string (AC-08)
- [ ] `UnloadFileAsync` clears state
- [ ] Invalid file paths return false without exception

---

## TASK-010-02-03: Implement DwgFileService.ExtractTableData() — read TableEntity cells

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` |
| Estimate | M |
| Depends On | TASK-010-02-02 |
| Blocks | TASK-010-02-04 |

### What to do
- Implement `ExtractTableDataAsync(layoutName)`:
  1. Find layout by name in `_document.Layouts`
  2. Iterate entities in layout's associated block
  3. Find `TableEntity` objects (ACadSharp type)
  4. For each table with 9 columns:
     - Read cell(0, 7) for "HEADER_ID" validation
     - Read cell(0, 8) for header_id value
     - Read rows 2+ for data (skip row 0=title, row 1=labels)
     - Map columns: position(0), product_no(1), width(2), height(3), len(4), thickness(5), qty(6), desc(7), detail_id(8)
  5. Filter empty rows (both qty and product_no empty)
  6. Strip MText formatting codes
- Wrap entire operation in try-catch — return empty list on ACadSharp failures
- Log warnings for skipped tables or parse errors

### How to verify
- [ ] Table data extracted with correct 9-column mapping (AC-03)
- [ ] Header validation checks column 7 for "HEADER_ID" text
- [ ] Empty rows are filtered (matching COM behavior)
- [ ] ACadSharp failures return empty list (AC-09)

---

## TASK-010-02-04: Implement DwgFileService.GetHeaderIds() and GetPRNumber()

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` |
| Estimate | S |
| Depends On | TASK-010-02-03 |
| Blocks | — |

### What to do
- Implement `GetHeaderIdsAsync()`:
  1. Iterate all layouts (excluding "Model")
  2. For each layout, find 9-column tables
  3. Read cell(0, 8) — the header_id value
  4. Return list of non-empty header IDs
- Implement `GetPRNumberAsync()`:
  1. Find attribute block in first layout (same logic as existing `GetBlockAttributes`)
  2. Look for `pr_no` attribute tag
  3. Return the value, or empty string if not found
- Both methods reuse internal helpers from `ExtractTableDataAsync`

### How to verify
- [ ] Header IDs extracted from cell(0,8) of each layout's table (AC-05)
- [ ] PR number extracted from block attributes (AC-06)
- [ ] Empty values handled gracefully

---

## TASK-010-02-05: Implement FileDrawingDataService — read operations delegation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` |
| Estimate | M |
| Depends On | TASK-010-02-01, TASK-010-01-02 |
| Blocks | TASK-010-03-02 |

### What to do
- Create `FileDrawingDataService` implementing `IDrawingDataService`
- Inject `IDwgFileService` and `SidecarIdStore` via constructor
- Implement read-only methods by delegating to `IDwgFileService`:
  ```csharp
  public AutoCADOperationMode Mode => AutoCADOperationMode.File;
  public bool IsReady => _dwgFileService.IsFileLoaded;
  public string? CurrentSource => _dwgFileService.LoadedFilePath;

  public Task<bool> ConnectOrLoadAsync(string? filePath)
      => _dwgFileService.LoadFileAsync(filePath ?? throw new ArgumentNullException());

  public Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync()
      => _dwgFileService.GetLayoutsAsync();  // from IDwgReaderService base

  public Task<LayoutData> ExtractParametersAsync(string layoutName)
      => /* combine attribute + table data */;
  ```
- `GetStatusAsync()` returns synthetic `AutoCADStatus` with file info instead of COM info
- `DisconnectOrUnloadAsync()` delegates to `UnloadFileAsync()`
- Write methods: defer to TASK-010-03-02

### How to verify
- [ ] All read methods delegate correctly to IDwgFileService
- [ ] Mode returns `File` (AC-07)
- [ ] IsReady reflects file loaded state (AC-01)
- [ ] Layouts exclude "Model" (AC-02)

---

## TASK-010-02-06: Add file picker and DWG info display to AutoCAD page

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | M |
| Depends On | TASK-010-01-07 |
| Blocks | — |

### What to do
- Add file mode panel (visible only when Mode == File):
  ```xml
  <StackPanel Visibility="{Binding IsFileMode, Converter={StaticResource BoolToVisibilityConverter}}">
    <GroupBox Header="DWG 檔案">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="{Binding CurrentSource}" TextTrimming="CharacterEllipsis"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button Content="開啟 DWG..." Command="{Binding OpenFileCommand}"/>
          <Button Content="關閉" Command="{Binding UnloadFileCommand}" Margin="8,0,0,0"/>
        </StackPanel>
      </Grid>
      <TextBlock Text="{Binding FileVersionInfo}" Foreground="Gray" FontSize="11"/>
    </GroupBox>
  </StackPanel>
  ```
- In AutoCADViewModel, add `OpenFileCommand`:
  1. Show `OpenFileDialog` with `.dwg` filter
  2. Call `_drawingDataService.ConnectOrLoadAsync(selectedPath)`
  3. Update status properties
- Add `IsFileMode` computed property
- Hide COM-specific "Connect" button when in file mode

### How to verify
- [ ] File picker shows only .dwg files (AC-01)
- [ ] DWG info displayed after loading (AC-08)
- [ ] Connect button hidden in file mode
- [ ] Open DWG button hidden in COM mode

---

## TASK-010-02-07: Refactor AutoCADViewModel to use IDrawingDataService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-010-02-05, TASK-010-01-02 |
| Blocks | — |

### What to do
- Replace direct `IGUIProxy` calls with `IDrawingDataService` calls:
  - `ConnectAsync()` → `_drawingDataService.ConnectOrLoadAsync()`
  - `DisconnectAsync()` → `_drawingDataService.DisconnectOrUnloadAsync()`
  - `GetStatusAsync()` → `_drawingDataService.GetStatusAsync()`
  - `GetLayouts()` → `_drawingDataService.GetLayoutsAsync()`
  - `ExtractParameters()` → `_drawingDataService.ExtractParametersAsync(layoutName)`
- Unify the existing if/else dual-path (around line 835) into single IDrawingDataService calls
- Add `IDrawingDataService` to constructor injection
- Keep `IGUIProxy` for now (other uses), but AutoCAD-specific calls go through `IDrawingDataService`

### How to verify
- [ ] No direct IGUIProxy AutoCAD calls remain in ViewModel (AC-07)
- [ ] COM mode behavior unchanged when using IDrawingDataService
- [ ] File mode works through same code paths

---

## TASK-010-02-08: Tests — DwgFileService file info, table read, attribute extraction

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/DwgFileServiceTests.cs` |
| Estimate | M |
| Depends On | TASK-010-02-02, TASK-010-02-03, TASK-010-02-04 |
| Blocks | — |

### What to do
- Create `DwgFileServiceTests.cs` with:
  ```csharp
  [Fact] LoadFile_ValidDwg_SetsIsFileLoaded()
  [Fact] LoadFile_NonExistentFile_ReturnsFalse()
  [Fact] LoadFile_InvalidFile_ReturnsFalse()
  [Fact] UnloadFile_ClearsState()
  [Fact] DwgVersion_ReturnsVersionString()
  [Fact] ExtractTableData_ValidTable_Returns9ColumnData()
  [Fact] ExtractTableData_InvalidTable_ReturnsEmpty()
  [Fact] GetHeaderIds_ReturnsHeadersFromTables()
  [Fact] GetPRNumber_ExtractsFromAttributes()
  ```
- For tests requiring actual DWG files, use a small test fixture DWG or mock ACadSharp types
- Table extraction tests should verify column mapping and row filtering

### How to verify
- [ ] All 9+ tests pass
- [ ] Table read tests verify 9-column mapping (AC-03)
- [ ] Error cases return empty results (AC-09)

---

## Dependency Graph
```
TASK-010-01-01 (Submodule)
    |
    +---> TASK-010-02-01 (IDwgFileService interface)
    |         |
    |         +---> TASK-010-02-02 (DwgFileService file management)
    |         |         |
    |         |         +---> TASK-010-02-03 (Table extraction)
    |         |         |         |
    |         |         |         +---> TASK-010-02-04 (Header IDs + PR number)
    |         |         |
    |         |         +---> TASK-010-02-08 (Tests)
    |         |
    |         +---> TASK-010-02-05 (FileDrawingDataService read)
    |
    +---> TASK-010-02-06 (AutoCAD page file picker)
    |
    +---> TASK-010-02-07 (AutoCADViewModel refactor)
```
