# FR-010: Dual-Mode AutoCAD Support (COM + ACadSharp)

> **Document Version**: 1.0
> **Last Updated**: 2026-02-13
> **Status**: Not Started
> **Priority**: P1

## 1. Overview

AutoCAD LT does not support the COM API, which means the current COM-only architecture cannot serve AutoCAD LT users. This feature introduces a dual-mode architecture that supports both full AutoCAD (COM real-time interaction) and AutoCAD LT (file-based mode via ACadSharp library).

A new unified interface `IDrawingDataService` abstracts the data access layer so that all ViewModels consume a single API regardless of whether the backend is COM or file-based. The Strategy pattern with a runtime dispatcher allows mode switching from the Settings page.

**Key Design Decisions**:
- ACadSharp is included as a **Git submodule** (not NuGet) to allow custom modifications to TableEntity/DwgWriter
- File mode uses **sidecar JSON** for table ID writeback (avoids DWG corruption risk)
- Block attribute writes use ACadSharp DXF export (stable path)
- Original DWG files are **never overwritten** in file mode

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-010-01 | CAD Engineer | Select between COM and File operation mode | I can use the application with AutoCAD LT or full AutoCAD |
| US-010-02 | CAD Engineer | Extract BOQ data and parameters from a DWG file without COM | I can process drawings from AutoCAD LT |
| US-010-03 | CAD Engineer | Write back IDs and attributes in file mode | My drawing data stays synchronized with Odoo even without COM |
| US-010-04 | System Admin | Switch between COM and File mode at runtime | I can adapt to different workstation configurations |

## 3. Python Reference

### Source Files
- `utility/util_autocad.py` (~1,301 lines) — All AutoCAD COM operations (COM mode reference)
- No Python equivalent for file mode — this is new functionality

### Key Patterns to Preserve
- Table structure: 9 columns, HEADER_ID in col 7, data rows start at row 2
- Attribute tag list: `product_name`, `spec`, `product_catelog`, `operation_flow`, `surface_treatment`, `color_name`, `color_no`
- Block identification: AcDbBlockReference with `project_name` or `job_working_plan_name` attribute tag
- Layout filtering: exclude "Model" layout

## 4. Functional Requirements

### Mode Selection

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-010-001 | Settings SHALL provide COM / File mode selection | Must |
| FR-010-002 | Selected mode SHALL persist across application restarts | Must |
| FR-010-003 | Mode change SHALL take effect immediately without restart | Must |
| FR-010-004 | COM mode SHALL remain the default for backwards compatibility | Must |

### Unified Data Interface

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-010-005 | All ViewModels SHALL consume `IDrawingDataService` instead of direct IGUIProxy/IAutoCADService calls | Must |
| FR-010-006 | `IDrawingDataService` SHALL expose `Mode`, `IsReady`, `CurrentSource` properties | Must |
| FR-010-007 | COM backend SHALL delegate to existing IAutoCADService + IGUIProxy | Must |
| FR-010-008 | File backend SHALL delegate to IDwgFileService (ACadSharp) | Must |
| FR-010-009 | Runtime dispatcher SHALL route calls based on current mode setting | Must |

### File Mode — Read Operations

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-010-010 | File mode SHALL read layouts from DWG via ACadSharp | Must |
| FR-010-011 | File mode SHALL extract table data (9-column tables) from DWG | Must |
| FR-010-012 | File mode SHALL extract block attributes from DWG | Must |
| FR-010-013 | File mode SHALL extract header IDs from table cells | Must |
| FR-010-014 | File mode SHALL extract PR number from block references | Must |
| FR-010-015 | File mode SHALL support loading DWG files via file picker | Must |

### File Mode — Write Operations

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-010-016 | Block attribute writes SHALL use ACadSharp DXF export | Must |
| FR-010-017 | Table ID writeback SHALL use sidecar JSON as primary strategy | Must |
| FR-010-018 | Sidecar JSON SHALL be saved alongside the DWG file (`{name}.boq-ids.json`) | Must |
| FR-010-019 | Original DWG files SHALL never be overwritten | Must |
| FR-010-020 | `IDrawingDataService.SupportsWrite` SHALL indicate write capability | Must |
| FR-010-021 | `IDrawingDataService.RecommendedWriteStrategy` SHALL guide UI behavior | Must |

### UI Integration

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-010-022 | AutoCAD page SHALL show mode indicator (COM / File) | Must |
| FR-010-023 | AutoCAD page SHALL provide file picker when in File mode | Must |
| FR-010-024 | BOQ page SHALL show info banner explaining sidecar writeback in file mode | Should |
| FR-010-025 | Settings > AutoCAD tab SHALL contain mode selection RadioButtons | Must |

## 5. UI Wireframe Description

### Settings > AutoCAD Tab (Mode Selection)
```
+--------------------------------------------------------------+
|  AutoCAD 設定                                                 |
+--------------------------------------------------------------+
|                                                               |
|  操作模式                                                     |
|  +----------------------------------------------------------+|
|  | (*) COM 模式 (完整版 AutoCAD — 即時連線)                    ||
|  | ( ) 檔案模式 (AutoCAD LT — 開啟 DWG 檔案)                  ||
|  +----------------------------------------------------------+|
|                                                               |
|  連線超時: [30] 秒                                            |
|  重試次數: [5]                                                |
|                                                               |
+--------------------------------------------------------------+
```

### AutoCAD Page (File Mode)
```
+--------------------------------------------------------------+
|                    AutoCAD Integration                        |
|                    [Mode: File] ℹ                             |
+--------------------------------------------------------------+
|                                                               |
|  DWG File                                                    |
|  +----------------------------------------------------------+|
|  | File: C:\Projects\Steel\Frame-001.dwg                    ||
|  | Version: AutoCAD 2018 (AC1032)                            ||
|  | Layouts: 4                                                ||
|  | [Open DWG...]  [Unload]                                   ||
|  +----------------------------------------------------------+|
|                                                               |
|  Layouts / Layout Details / Table Data                       |
|  (same as COM mode)                                          |
|                                                               |
+--------------------------------------------------------------+
```

### BOQ Page (File Mode Banner)
```
+--------------------------------------------------------------+
|  ℹ 檔案模式：ID 回寫將儲存至 sidecar JSON 檔案                  |
|    ({drawing}.boq-ids.json)，不會修改原始 DWG。                 |
+--------------------------------------------------------------+
```

## 6. Data Model

### Core Enums and Interface

```csharp
public enum AutoCADOperationMode { COM, File }
public enum WriteStrategy { DirectDwg, ExportDxf, SidecarJson, WriteUnsupported }

public interface IDrawingDataService
{
    AutoCADOperationMode Mode { get; }
    bool IsReady { get; }
    string? CurrentSource { get; }

    // Connect / Load
    Task<bool> ConnectOrLoadAsync(string? filePath = null);
    Task DisconnectOrUnloadAsync();
    Task<AutoCADStatus> GetStatusAsync();

    // READ
    Task<IReadOnlyList<LayoutInfo>> GetLayoutsAsync();
    Task<LayoutData> ExtractParametersAsync(string layoutName);
    Task<IReadOnlyList<string>> GetHeaderIdsAsync();
    Task<string> GetPRNumberAsync();
    Task<Dictionary<string, string>> GetAttributeBlockAsync(string? layoutName = null);

    // WRITE
    Task<WritebackResult> WriteTableIdsAsync(string layoutName, string headerId, IList<WritebackDetail> details);
    Task<List<string>> SetAttributeValuesAsync(Dictionary<string, string> attributes, string? layoutName = null);
    bool SupportsWrite { get; }
    WriteStrategy RecommendedWriteStrategy { get; }
    Task<bool> SaveAsync();
}
```

### Sidecar JSON Schema

```json
{
  "version": 1,
  "source_dwg": "drawing.dwg",
  "modified_at": "2026-02-13T10:30:00Z",
  "layouts": [{
    "name": "Layout1",
    "header_id": "BOQ-001-H",
    "details": [{ "product_no": "ST-001", "detail_id": "BOQ-001-D-01" }]
  }],
  "attributes": { "product_name": "Steel Plate", "spec": "SS400" }
}
```

## 7. API/Service Dependencies

| Service | Interface | Methods Used |
|---------|-----------|-------------|
| Drawing Data Service | `IDrawingDataService` | All 12 methods — unified interface for all ViewModels |
| COM Backend | `IAutoCADService` + `IGUIProxy` | Existing COM operations (unchanged) |
| File Backend | `IDwgFileService` | ACadSharp read/write operations |
| Sidecar Store | `SidecarIdStore` | JSON sidecar read/write for file mode |
| Settings Service | `ISettingsService` | Mode persistence, Swagger URL |
| DWG Reader | `IDwgReaderService` | Existing read-only operations (extended by IDwgFileService) |

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-010-001 | COM mode requires AutoCAD running; File mode requires DWG file loaded |
| VR-010-002 | File picker only shows .dwg files |
| VR-010-003 | Sidecar JSON filename must match DWG filename |
| VR-010-004 | Mode switch clears current connection/loaded file state |
| VR-010-005 | Write operations check `SupportsWrite` before attempting |
| VR-010-006 | File mode write operations check file is not read-only |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| DWG file not found | "找不到 DWG 檔案: {path}" | Show error, prompt file picker |
| DWG file corrupted | "無法讀取 DWG 檔案: {details}" | Show error with details |
| ACadSharp table read failure | "讀取表格資料時發生錯誤（佈局 '{name}'）" | Skip table, return empty |
| Sidecar JSON write failure | "無法儲存 ID 回寫檔案: {details}" | Show error, suggest manual save |
| DXF export failure | "DXF 匯出失敗: {details}" | Show error, suggest retry |
| Mode switch while busy | "請等待目前操作完成後再切換模式。" | Disable mode switch during operations |
| File locked by another process | "DWG 檔案被其他程式鎖定: {path}" | Show error, suggest close other app |

## 10. Implementation Notes

### C# Target Files

**New Files**:
- `OdooAutoCAD.Core/AutoCAD/IDrawingDataService.cs` — Unified interface
- `OdooAutoCAD.Core/AutoCAD/IDwgFileService.cs` — Extends IDwgReaderService + write
- `OdooAutoCAD.Core/AutoCAD/DwgFileService.cs` — ACadSharp implementation
- `OdooAutoCAD.Core/AutoCAD/ComDrawingDataService.cs` — COM backend
- `OdooAutoCAD.Core/AutoCAD/FileDrawingDataService.cs` — File backend
- `OdooAutoCAD.Core/AutoCAD/DrawingDataServiceDispatcher.cs` — Runtime mode switch
- `OdooAutoCAD.Core/AutoCAD/SidecarIdStore.cs` — JSON sidecar

**Modified Files**:
- `OdooAutoCAD.Core.csproj` — NuGet → ProjectReference (ACadSharp submodule)
- `ViewModels/BOQViewModel.cs` — IGUIProxy → IDrawingDataService
- `ViewModels/PurchaseRequisitionViewModel.cs` — IGUIProxy → IDrawingDataService
- `ViewModels/ParameterConfigViewModel.cs` — IGUIProxy → IDrawingDataService
- `ViewModels/AutoCADViewModel.cs` — Unified dual-path → IDrawingDataService
- `ViewModels/DashboardViewModel.cs` — Mode-aware connect/load
- `App.xaml.cs` — Register new DI services
- `ViewModels/SettingsViewModel.cs` — Add AutoCADOperationMode property
- `Views/Pages/SettingsPage.xaml` — Mode selection RadioButtons
- `Views/Pages/AutoCADPage.xaml` — Mode indicator + file picker
- `Views/Pages/BOQPage.xaml` — File mode info banner

### Architecture Pattern
```
IDrawingDataService (Strategy interface)
├── ComDrawingDataService  → IAutoCADService + IGUIProxy (existing COM)
├── FileDrawingDataService → IDwgFileService (ACadSharp read/write)
└── DrawingDataServiceDispatcher (runtime mode switch)
```

### Write Strategy (File Mode)

| Operation | Strategy | Reason |
|-----------|----------|--------|
| Read layouts/attributes/tables | ACadSharp DwgReader | Stable |
| Write block attributes | ACadSharp DXF export | DxfWriter supports all versions |
| Write table IDs | Sidecar JSON (primary) | TableEntity is WIP in ACadSharp |
| Write table IDs | DXF export (secondary) | If TableEntity becomes stable |

### ACadSharp Submodule
- Path: `csharp/libs/ACadSharp/`
- Referenced as ProjectReference from OdooAutoCAD.Core.csproj
- .NET Standard 2.0, compatible with .NET 8
- Allows custom TableEntity/DwgWriter modifications if needed

### Key Differences from COM Mode
- COM: real-time connection to running AutoCAD instance
- File: load DWG from disk, process offline, save as DXF or sidecar JSON
- COM: IGUIProxy ensures STA thread safety
- File: no COM threading concerns, can run on any thread
- COM: write directly to AutoCAD drawing
- File: write via DXF export or sidecar JSON (no DWG overwrite)
