# TASKS: US-010-04 — Runtime Mode Switching

> **Parent US**: [US-010-04](US-010-04-mode-switch-runtime.md)
> **Parent FR**: [FR-010](FR-010-dual-mode-autocad.md)
> **Priority**: P1
> **Tasks**: 9 | **Effort**: 2S + 7M
> **Status**: Not Started

## Prerequisites
- [ ] US-010-01 complete (IDrawingDataService, Dispatcher, DI)
- [ ] US-010-02 complete (FileDrawingDataService read)
- [ ] US-010-03 complete (FileDrawingDataService write, SidecarIdStore)
- [ ] Existing ViewModels use IGUIProxy for AutoCAD operations

## Acceptance Criteria
- [ ] AC-01: Settings mode change immediately updates DrawingDataServiceDispatcher
- [ ] AC-02: Mode switch disconnects/unloads before switching
- [ ] AC-03: All ViewModels reflect new mode without page navigation
- [ ] AC-04: Dashboard buttons adapt to current mode
- [ ] AC-05: Existing tests pass in both modes (mocked)
- [ ] AC-06: BOQViewModel uses IDrawingDataService
- [ ] AC-07: PurchaseRequisitionViewModel uses IDrawingDataService
- [ ] AC-08: ParameterConfigViewModel uses IDrawingDataService
- [ ] AC-09: DashboardViewModel uses IDrawingDataService

---

## TASK-010-04-01: Implement ComDrawingDataService wrapping IAutoCADService + IGUIProxy

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/ComDrawingDataService.cs` |
| Estimate | M |
| Depends On | TASK-010-01-02 |
| Blocks | TASK-010-04-02, TASK-010-04-03, TASK-010-04-04, TASK-010-04-05 |

### What to do
- Create `ComDrawingDataService` implementing `IDrawingDataService`
- Inject `IAutoCADService` and `IGUIProxy` via constructor
- Delegate all methods to existing COM infrastructure:
  ```csharp
  public AutoCADOperationMode Mode => AutoCADOperationMode.COM;
  public bool IsReady => _autoCADService.IsConnected;
  public string? CurrentSource => _autoCADService.CurrentDocumentName;
  public bool SupportsWrite => true;
  public WriteStrategy RecommendedWriteStrategy => WriteStrategy.DirectDwg;

  public async Task<bool> ConnectOrLoadAsync(string? filePath = null)
  {
      var response = await _guiProxy.ExecuteInGuiAsync("autocad_connect");
      return response.Success && response.Result is bool b && b;
  }

  public async Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync()
  {
      var response = await _guiProxy.ExecuteInGuiAsync("autocad_get_layouts");
      return (response.Result as IReadOnlyList<LayoutInfo>) ?? Array.Empty<LayoutInfo>();
  }

  // ... similar delegation for all other methods
  ```
- Map each `IDrawingDataService` method to the corresponding GUIProxy handler:
  - `ConnectOrLoadAsync` → `"autocad_connect"`
  - `DisconnectOrUnloadAsync` → `"autocad_disconnect"`
  - `GetStatusAsync` → `"autocad_get_status"`
  - `GetLayoutsAsync` → `"autocad_get_layouts"`
  - `ExtractParametersAsync` → `"autocad_get_layout_values"`
  - `GetHeaderIdsAsync` → `"autocad_get_header_ids"`
  - `GetPRNumberAsync` → `"autocad_get_pr_number"`
  - `GetAttributeBlockAsync` → `"autocad_get_attribute_block"`
  - `WriteTableIdsAsync` → `"autocad_write_table_ids"`
  - `SetAttributeValuesAsync` → `"autocad_set_attribute_values"`
  - `SaveAsync` → return true (COM saves immediately)

### How to verify
- [ ] All IDrawingDataService methods delegate to correct GUIProxy handlers
- [ ] Mode returns COM
- [ ] SupportsWrite returns true
- [ ] RecommendedWriteStrategy returns DirectDwg
- [ ] COM behavior unchanged from previous implementation

---

## TASK-010-04-02: Refactor BOQViewModel → IDrawingDataService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQViewModel.cs` |
| Estimate | M |
| Depends On | TASK-010-04-01 |
| Blocks | TASK-010-04-06 |

### What to do
- Add `IDrawingDataService _drawingDataService` to constructor (alongside existing dependencies)
- Replace all `_guiProxy.ExecuteInGuiAsync("autocad_*")` calls with `_drawingDataService` methods:
  - BOQ extraction: `_guiProxy.ExecuteInGuiAsync("autocad_get_layout_values")` → `_drawingDataService.ExtractParametersAsync(layoutName)`
  - ID writeback: `_guiProxy.ExecuteInGuiAsync("autocad_write_table_ids")` → `_drawingDataService.WriteTableIdsAsync(layout, headerId, details)`
  - Layout listing: use `_drawingDataService.GetLayoutsAsync()`
- Add `IsFileMode` property for UI binding
- Add `IsReady` check before operations: `_drawingDataService.IsReady`
- Keep `IOdooService`, `ISettingsService` dependencies unchanged

### How to verify
- [ ] No direct `_guiProxy.ExecuteInGuiAsync("autocad_*")` calls remain (AC-06)
- [ ] BOQ extraction works through IDrawingDataService
- [ ] ID writeback works through IDrawingDataService
- [ ] COM mode behavior identical to before

---

## TASK-010-04-03: Refactor PurchaseRequisitionViewModel → IDrawingDataService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-010-04-01 |
| Blocks | TASK-010-04-07 |

### What to do
- Add `IDrawingDataService _drawingDataService` to constructor
- Replace header ID extraction:
  - `_guiProxy.ExecuteInGuiAsync("autocad_get_header_ids")` → `_drawingDataService.GetHeaderIdsAsync()`
- Replace PR number extraction if used:
  - `_drawingDataService.GetPRNumberAsync()`
- Add `IsReady` check before operations
- Keep `IOdooService` dependency unchanged

### How to verify
- [ ] Header ID extraction uses IDrawingDataService (AC-07)
- [ ] No direct `_guiProxy.ExecuteInGuiAsync("autocad_*")` calls remain
- [ ] COM mode behavior unchanged

---

## TASK-010-04-04: Refactor ParameterConfigViewModel → IDrawingDataService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/ParameterConfigViewModel.cs` |
| Estimate | M |
| Depends On | TASK-010-04-01 |
| Blocks | TASK-010-04-08 |

### What to do
- Add `IDrawingDataService _drawingDataService` to constructor
- Replace attribute read/write operations:
  - `_guiProxy.ExecuteInGuiAsync("autocad_get_attribute_block")` → `_drawingDataService.GetAttributeBlockAsync()`
  - `_guiProxy.ExecuteInGuiAsync("autocad_set_attribute_values")` → `_drawingDataService.SetAttributeValuesAsync(attributes)`
- After attribute write in file mode, prompt user to save:
  ```csharp
  if (_drawingDataService.Mode == AutoCADOperationMode.File)
  {
      await _drawingDataService.SaveAsync(); // DXF export
      StatusMessage = "屬性已儲存至 DXF 檔案";
  }
  ```
- Add `IsReady` check before operations

### How to verify
- [ ] Attribute read/write uses IDrawingDataService (AC-08)
- [ ] No direct `_guiProxy.ExecuteInGuiAsync("autocad_*")` calls remain
- [ ] File mode triggers DXF export after attribute write
- [ ] COM mode behavior unchanged

---

## TASK-010-04-05: Refactor DashboardViewModel → mode-aware connect/load

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/DashboardViewModel.cs` |
| Estimate | M |
| Depends On | TASK-010-04-01 |
| Blocks | — |

### What to do
- Add `IDrawingDataService _drawingDataService` to constructor
- Replace AutoCAD connect logic:
  - `_guiProxy.ExecuteInGuiAsync("autocad_connect")` → `_drawingDataService.ConnectOrLoadAsync()`
- For file mode, the connect button should open a file picker:
  ```csharp
  [RelayCommand]
  private async Task ConnectAutoCADAsync()
  {
      if (_drawingDataService.Mode == AutoCADOperationMode.File)
      {
          // File picker logic (OpenFileDialog for .dwg)
          var dialog = new OpenFileDialog { Filter = "DWG Files (*.dwg)|*.dwg" };
          if (dialog.ShowDialog() == true)
              await _drawingDataService.ConnectOrLoadAsync(dialog.FileName);
      }
      else
      {
          await _drawingDataService.ConnectOrLoadAsync();
      }
  }
  ```
- Update button text: "連接 AutoCAD" in COM mode, "開啟 DWG..." in file mode
- Update status indicators based on `_drawingDataService.IsReady`
- Subscribe to `DrawingDataServiceDispatcher.ModeChanged` to refresh UI

### How to verify
- [ ] Dashboard adapts to current mode (AC-04, AC-09)
- [ ] COM mode: "連接 AutoCAD" button behavior unchanged
- [ ] File mode: "開啟 DWG..." button opens file picker
- [ ] Status indicators reflect mode-appropriate state

---

## TASK-010-04-06: Update BOQViewModelTests for dual-mode

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/ViewModels/BOQViewModelTests.cs` |
| Estimate | M |
| Depends On | TASK-010-04-02 |
| Blocks | — |

### What to do
- Replace `Mock<IGUIProxy>` with `Mock<IDrawingDataService>` for AutoCAD operations
- Update existing test setup to inject `IDrawingDataService` mock
- Add new tests:
  ```csharp
  [Fact] ExtractBOQ_FileMode_UsesDrawingDataService()
  [Fact] WritebackIds_FileMode_UsesSidecarStrategy()
  [Fact] WritebackIds_ComMode_UsesDirectStrategy()
  [Fact] IsFileMode_ReturnsTrue_WhenFileModeActive()
  ```
- Keep existing test logic — only change the mock source
- Verify all 15+ existing BOQ tests still pass with new mock

### How to verify
- [ ] All existing BOQ tests pass with IDrawingDataService mock (AC-05)
- [ ] New dual-mode tests pass
- [ ] Test coverage maintained

---

## TASK-010-04-07: Update PurchaseRequisitionViewModelTests for dual-mode

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/ViewModels/PurchaseRequisitionViewModelTests.cs` |
| Estimate | S |
| Depends On | TASK-010-04-03 |
| Blocks | — |

### What to do
- Replace `Mock<IGUIProxy>` with `Mock<IDrawingDataService>` for header ID extraction
- Update test setup to inject `IDrawingDataService` mock
- Add:
  ```csharp
  [Fact] ConvertBOQToPR_FileMode_GetsHeaderIdsFromDrawingDataService()
  ```
- Keep existing test logic intact
- Verify all 18 existing PR tests still pass

### How to verify
- [ ] All existing PR tests pass (AC-05)
- [ ] Header ID extraction mocked through IDrawingDataService (AC-07)

---

## TASK-010-04-08: Update ParameterConfigViewModelTests for dual-mode

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/ViewModels/ParameterConfigViewModelTests.cs` |
| Estimate | S |
| Depends On | TASK-010-04-04 |
| Blocks | — |

### What to do
- Replace `Mock<IGUIProxy>` with `Mock<IDrawingDataService>` for attribute operations
- Update test setup to inject `IDrawingDataService` mock
- Add:
  ```csharp
  [Fact] Submit_FileMode_TriggersAttributeWriteAndSave()
  [Fact] Submit_ComMode_WritesAttributesDirectly()
  ```
- Keep existing test logic intact
- Verify all 16 existing ParameterConfig tests still pass

### How to verify
- [ ] All existing tests pass (AC-05)
- [ ] Attribute operations mocked through IDrawingDataService (AC-08)

---

## TASK-010-04-09: Tests — ComDrawingDataService delegation

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/DrawingDataServiceTests.cs` |
| Estimate | M |
| Depends On | TASK-010-04-01 |
| Blocks | — |

### What to do
- Add to `DrawingDataServiceTests.cs` (or create new section):
  ```csharp
  [Fact] ComBackend_ConnectOrLoadAsync_DelegatesToGUIProxy()
  [Fact] ComBackend_GetLayoutsAsync_DelegatesToGUIProxy()
  [Fact] ComBackend_ExtractParametersAsync_DelegatesToGUIProxy()
  [Fact] ComBackend_WriteTableIdsAsync_DelegatesToGUIProxy()
  [Fact] ComBackend_SetAttributeValuesAsync_DelegatesToGUIProxy()
  [Fact] ComBackend_Mode_ReturnsCOM()
  [Fact] ComBackend_SupportsWrite_ReturnsTrue()
  [Fact] ComBackend_SaveAsync_ReturnsTrue()
  ```
- Mock `IGUIProxy` and verify correct handler names are called
- Verify response parsing matches existing ViewModel expectations

### How to verify
- [ ] All 8 tests pass
- [ ] Each method verified to delegate to correct GUIProxy handler
- [ ] Response parsing matches expected types

---

## Dependency Graph
```
TASK-010-04-01 (ComDrawingDataService)
    |
    +---> TASK-010-04-02 (BOQViewModel refactor)
    |         |
    |         +---> TASK-010-04-06 (BOQ tests)
    |
    +---> TASK-010-04-03 (PRViewModel refactor)
    |         |
    |         +---> TASK-010-04-07 (PR tests)
    |
    +---> TASK-010-04-04 (ParameterConfigViewModel refactor)
    |         |
    |         +---> TASK-010-04-08 (ParameterConfig tests)
    |
    +---> TASK-010-04-05 (DashboardViewModel refactor)
    |
    +---> TASK-010-04-09 (ComDrawingDataService tests)
```
