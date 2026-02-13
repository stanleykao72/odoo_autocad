# TASKS: US-010-01 — Select Operation Mode

> **Parent US**: [US-010-01](US-010-01-select-operation-mode.md)
> **Parent FR**: [FR-010](FR-010-dual-mode-autocad.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 4S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] Sprint 2 infrastructure complete (SettingsViewModel, ISettingsService)
- [ ] Sprint 3 infrastructure complete (IAutoCADService, IGUIProxy)
- [ ] ACadSharp submodule added and NuGet reference replaced

## Acceptance Criteria
- [ ] AC-01: Settings > AutoCAD tab displays two RadioButtons: "COM 模式" and "檔案模式"
- [ ] AC-02: COM mode is selected by default on fresh install
- [ ] AC-03: Selected mode persists across application restarts via ISettingsService
- [ ] AC-04: Mode change triggers DrawingDataServiceDispatcher to switch backend
- [ ] AC-05: SettingsViewModel exposes AutoCADOperationMode property with two-way binding
- [ ] AC-06: AutoCAD page shows current mode indicator
- [ ] AC-07: Mode descriptions explain the difference

---

## TASK-010-01-01: Add ACadSharp git submodule and replace NuGet reference

| Field | Value |
|-------|-------|
| Target | `csharp/libs/ACadSharp/`, `OdooAutoCAD.Core.csproj` |
| Estimate | M |
| Depends On | — |
| Blocks | TASK-010-01-02, TASK-010-02-02 |

### What to do
- Add git submodule: `git submodule add https://github.com/DomCR/ACadSharp.git csharp/libs/ACadSharp`
- In `OdooAutoCAD.Core.csproj`, remove `<PackageReference Include="ACadSharp" Version="..." />`
- Add `<ProjectReference Include="..\..\libs\ACadSharp\src\ACadSharp\ACadSharp.csproj" />`
- Run `dotnet build OdooAutoCADIntegration.sln` to verify 0 errors
- Update `.gitmodules` if needed

### How to verify
- [ ] Submodule exists at `csharp/libs/ACadSharp/` with valid `.git` reference
- [ ] `OdooAutoCAD.Core.csproj` has ProjectReference, not PackageReference for ACadSharp
- [ ] Solution builds with 0 errors

---

## TASK-010-01-02: Define AutoCADOperationMode, WriteStrategy enums and IDrawingDataService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/IDrawingDataService.cs` |
| Estimate | M |
| Depends On | TASK-010-01-01 |
| Blocks | TASK-010-01-05, TASK-010-02-05, TASK-010-04-01 |

### What to do
- Create new file `IDrawingDataService.cs` in `OdooAutoCAD.Core/AutoCAD/`
- Define `AutoCADOperationMode` enum: `COM`, `File`
- Define `WriteStrategy` enum: `DirectDwg`, `ExportDxf`, `SidecarJson`, `WriteUnsupported`
- Define `WritebackResult` record: `Success`, `Message`, `Strategy`
- Define `IDrawingDataService` interface with ~12 methods:
  ```csharp
  AutoCADOperationMode Mode { get; }
  bool IsReady { get; }
  string? CurrentSource { get; }
  Task<bool> ConnectOrLoadAsync(string? filePath = null);
  Task DisconnectOrUnloadAsync();
  Task<AutoCADStatus> GetStatusAsync();
  Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync();
  Task<LayoutData> ExtractParametersAsync(string layoutName);
  Task<IReadOnlyList<string>> GetHeaderIdsAsync();
  Task<string> GetPRNumberAsync();
  Task<Dictionary<string, string>> GetAttributeBlockAsync(string? layoutName = null);
  Task<WritebackResult> WriteTableIdsAsync(string layoutName, string headerId, IList<WritebackDetail> details);
  Task<List<string>> SetAttributeValuesAsync(Dictionary<string, string> attributes, string? layoutName = null);
  bool SupportsWrite { get; }
  WriteStrategy RecommendedWriteStrategy { get; }
  Task<bool> SaveAsync();
  ```
- Define `WritebackDetail` record: `ProductNo`, `DetailId`
- Use existing types: `AutoCADStatus`, `LayoutInfo`, `LayoutData`

### How to verify
- [ ] `AutoCADOperationMode` enum exists with `COM` and `File` values (AC-01)
- [ ] `WriteStrategy` enum exists with 4 values
- [ ] `IDrawingDataService` interface compiles with all 12+ methods
- [ ] Solution builds without errors

---

## TASK-010-01-03: Add AutoCADOperationMode property to SettingsViewModel and ISettingsService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/SettingsViewModel.cs`, `src/OdooAutoCAD.Core/Services/ISettingsService.cs` |
| Estimate | S |
| Depends On | TASK-010-01-02 |
| Blocks | TASK-010-01-04, TASK-010-01-05 |

### What to do
- Add `AutoCADOperationMode` to `ISettingsService` as a gettable/settable preference
- Add to `SettingsService` implementation: store as UserPreference key `"AutoCADMode"` with values `"COM"` / `"File"`
- Default to `AutoCADOperationMode.COM` when key not found
- In `SettingsViewModel`, add:
  ```csharp
  [ObservableProperty]
  private AutoCADOperationMode _autoCADMode = AutoCADOperationMode.COM;
  ```
- Load mode in `LoadSettingsAsync()`, save in `SaveSettingsAsync()`
- Mark form dirty when mode changes

### How to verify
- [ ] `SettingsViewModel.AutoCADMode` is an observable property (AC-05)
- [ ] Default value is `COM` (AC-02)
- [ ] Value persists via ISettingsService (AC-03)

---

## TASK-010-01-04: Add mode selection RadioButtons to Settings > AutoCAD tab

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | TASK-010-01-03 |
| Blocks | — |

### What to do
- In the AutoCAD tab of SettingsPage.xaml, add a GroupBox "操作模式":
  ```xml
  <GroupBox Header="操作模式" Margin="0,0,0,16">
    <StackPanel>
      <RadioButton Content="COM 模式 (完整版 AutoCAD — 即時連線)"
                   IsChecked="{Binding AutoCADMode, Converter={StaticResource EnumBoolConverter}, ConverterParameter=COM}"
                   Margin="0,4"/>
      <TextBlock Text="適用於完整版 AutoCAD，透過 COM 即時操作圖面。"
                 Foreground="Gray" FontSize="11" Margin="20,0,0,8"/>
      <RadioButton Content="檔案模式 (AutoCAD LT — 開啟 DWG 檔案)"
                   IsChecked="{Binding AutoCADMode, Converter={StaticResource EnumBoolConverter}, ConverterParameter=File}"
                   Margin="0,4"/>
      <TextBlock Text="適用於 AutoCAD LT，直接讀取 DWG 檔案。ID 回寫使用 sidecar JSON。"
                 Foreground="Gray" FontSize="11" Margin="20,0,0,4"/>
    </StackPanel>
  </GroupBox>
  ```
- Add `EnumBoolConverter` if not already available (convert enum to bool for RadioButton binding)

### How to verify
- [ ] Two RadioButtons visible in Settings > AutoCAD tab (AC-01)
- [ ] Each option has descriptive text (AC-07)
- [ ] COM mode selected by default (AC-02)
- [ ] Selection changes bind to SettingsViewModel.AutoCADMode

---

## TASK-010-01-05: Implement DrawingDataServiceDispatcher with mode switching

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/DrawingDataServiceDispatcher.cs` |
| Estimate | L |
| Depends On | TASK-010-01-02, TASK-010-01-03 |
| Blocks | TASK-010-01-06 |

### What to do
- Create `DrawingDataServiceDispatcher` implementing `IDrawingDataService`
- Inject `ComDrawingDataService` and `FileDrawingDataService` via constructor
- Inject `ISettingsService` to read current mode
- Implement all interface methods by delegating to the active backend:
  ```csharp
  private IDrawingDataService ActiveBackend =>
      _settingsService.GetAutoCADMode() == AutoCADOperationMode.COM
          ? _comBackend
          : _fileBackend;
  ```
- Add `SwitchModeAsync(AutoCADOperationMode newMode)` method:
  1. Disconnect/unload current backend
  2. Update settings
  3. Fire `ModeChanged` event
- Expose `event Action<AutoCADOperationMode>? ModeChanged`
- `Mode` property delegates to active backend's mode

### How to verify
- [ ] All IDrawingDataService calls route to correct backend based on mode (AC-04)
- [ ] Mode switch disconnects current backend before switching
- [ ] ModeChanged event fires on switch

---

## TASK-010-01-06: Register new DI services in App.xaml.cs

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-010-01-05 |
| Blocks | — |

### What to do
- Register in ConfigureServices:
  ```csharp
  services.AddSingleton<ComDrawingDataService>();
  services.AddSingleton<FileDrawingDataService>();
  services.AddSingleton<DrawingDataServiceDispatcher>();
  services.AddSingleton<IDrawingDataService>(sp => sp.GetRequiredService<DrawingDataServiceDispatcher>());
  services.AddSingleton<IDwgFileService, DwgFileService>();
  services.AddSingleton<SidecarIdStore>();
  ```
- Ensure `ComDrawingDataService` can resolve `IAutoCADService` and `IGUIProxy`
- Ensure `FileDrawingDataService` can resolve `IDwgFileService` and `SidecarIdStore`

### How to verify
- [ ] All new services resolve from DI container without error
- [ ] Application starts successfully with new registrations

---

## TASK-010-01-07: Add mode indicator to AutoCAD page

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | TASK-010-01-02 |
| Blocks | — |

### What to do
- Add a mode indicator badge near the page title:
  ```xml
  <Border Background="{Binding Mode, Converter={StaticResource ModeToColorConverter}}"
          CornerRadius="4" Padding="8,2" Margin="8,0">
    <TextBlock Text="{Binding ModeDisplayText}" FontSize="11" Foreground="White"/>
  </Border>
  ```
- In AutoCADViewModel, add computed property:
  ```csharp
  public string ModeDisplayText => _drawingDataService.Mode == AutoCADOperationMode.COM
      ? "Mode: COM" : "Mode: File";
  ```
- Use simple color: Blue for COM, Green for File

### How to verify
- [ ] Mode badge visible on AutoCAD page (AC-06)
- [ ] Badge updates when mode changes

---

## TASK-010-01-08: Tests — Dispatcher mode switching and Settings persistence

| Field | Value |
|-------|-------|
| Target | `tests/OdooAutoCAD.Integration.Tests/Services/DrawingDataServiceTests.cs` |
| Estimate | M |
| Depends On | TASK-010-01-05, TASK-010-01-03 |
| Blocks | — |

### What to do
- Create `DrawingDataServiceTests.cs` with:
  ```csharp
  [Fact] SwitchMode_FromComToFile_DisconnectsComAndActivatesFile()
  [Fact] SwitchMode_FromFileToCom_UnloadsFileAndActivatesCom()
  [Fact] Mode_DefaultsCom_WhenSettingsNotConfigured()
  [Fact] GetLayoutsAsync_InComMode_DelegatesToComBackend()
  [Fact] GetLayoutsAsync_InFileMode_DelegatesToFileBackend()
  [Fact] ModeChanged_Event_FiredOnSwitch()
  ```
- Mock `ComDrawingDataService`, `FileDrawingDataService`, `ISettingsService`
- Verify delegation behavior and event firing

### How to verify
- [ ] All 6+ tests pass
- [ ] Dispatcher routes to correct backend based on mode (AC-04)
- [ ] Default mode is COM (AC-02)

---

## Dependency Graph
```
TASK-010-01-01 (Submodule + ProjectReference)
    |
    +---> TASK-010-01-02 (IDrawingDataService interface)
              |
              +---> TASK-010-01-03 (SettingsViewModel mode property)
              |         |
              |         +---> TASK-010-01-04 (Settings UI RadioButtons)
              |         |
              |         +---> TASK-010-01-05 (Dispatcher implementation)
              |                   |
              |                   +---> TASK-010-01-06 (DI registration)
              |                   |
              |                   +---> TASK-010-01-08 (Tests)
              |
              +---> TASK-010-01-07 (Mode indicator on AutoCAD page)
```
