# FR-002: AutoCAD Integration Page

> **Document Version**: 1.1
> **Last Updated**: 2026-02-11
> **Status**: In Progress (Sprint 3 Complete — US-002-01, US-002-02, US-002-03; COM threading fix applied)
> **Priority**: P1

## 1. Overview

The AutoCAD Integration Page manages the connection to AutoCAD via COM automation, displays drawing information, and provides controls for parameter extraction and layout operations. This is one of the most complex features, bridging COM interop with the WPF UI through thread-safe proxy patterns.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-002-01 | CAD Engineer | Connect to a running AutoCAD instance | I can extract parameters from my current drawing |
| US-002-02 | CAD Engineer | See all layouts in my drawing | I can select which layout to work with |
| US-002-03 | CAD Engineer | Extract parameters from a layout | I can map drawing data to Odoo products |
| US-002-04 | CAD Engineer | See the current drawing file name and path | I know which drawing is active |
| US-002-05 | CAD Engineer | View extracted PR number and project info | I can verify the correct project context |
| US-002-06 | CAD Engineer | Clear table IDs from layouts | I can re-extract data when needed |
| US-002-07 | System Admin | Monitor COM connection status | I can troubleshoot connectivity issues |

## 3. Python Reference

### Source Files
- `utility/util_autocad.py` (~1,301 lines) - All AutoCAD COM operations
- `forms/form_main_modern.py` - AutoCAD connection button, status display
- `forms/form_autocad_param.py` - Parameter input dialogs

### Key Functions
- `connect_autocad(main_body)` - COM connection with retry (5 attempts)
- `get_layouts_values()` - Extracts all layout data including tables
- `get_active_layout()` - Returns current active layout
- `get_doc_layouts()` - List all layouts excluding "Model"
- `process_pr_no(layout)` - Extracts PR number and project from layout
- `get_attribute_values(layout, block_name, tag_list)` - Gets block attributes
- `get_table_data(table_list)` - Parses table rows into detail dicts
- `set_attribute_value(block, tag, val)` - Updates AutoCAD block attributes
- `clear_table_id(layout)` / `clear_all_tables_id()` - Clear ID columns
- `scan_entities()` - Scan ModelSpace for entity listing
- `create_new_drawing()`, `draw_line()`, `draw_circle()`, `create_text()` - Drawing creation
- `add_dimension()` - Add dimension annotations
- `set_layer()`, `list_layers()` - Layer management
- `LM_UnFormat(s, mtx)` - Remove AutoCAD text formatting codes

### Connection Flow (Python)
1. `pythoncom.CoInitialize()` - Initialize COM threading
2. Try `client.GetActiveObject("AutoCAD.Application")` - Attach to running instance
3. Fallback: `client.Dispatch("AutoCAD.Application")` - Launch new instance
4. Set AutoCAD visible
5. Retry up to 5 times to get `ActiveDocument` (1-second intervals)
6. Extract file name and path
7. Call `process_pr_no()` to get project context

### Tag List (10 standard attributes extracted from layouts)
```
pr_no, project_name, job_working_plan_name, product_name,
product_catelog, spec, surface_treatment, operation_flow,
color_name, color_no
```

### Table Structure
- 9 columns: position, product_no, width, height, len, thickness, qty, desc, detail_id
- Validation: must have 9 columns and "HEADER_ID" in column 7 header
- Row 0 = header (contains header_id in col 8), Row 1 = labels (skipped), Row 2+ = data

## 4. Functional Requirements

### Connection Management

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-002-001 | Page SHALL provide a Connect button that connects to AutoCAD via COM | Must |
| FR-002-002 | Connection SHALL first try GetActiveObject, then fallback to Dispatch | Must |
| FR-002-003 | Page SHALL display connection status (Connected/Disconnected) with visual indicator | Must |
| FR-002-004 | Page SHALL display AutoCAD version and application info when connected | Should |
| FR-002-005 | Page SHALL display current document filename and full path | Must |
| FR-002-006 | Connect button SHALL change appearance when connected (e.g., show Disconnect) | Should |
| FR-002-007 | All COM operations SHALL execute on the GUI/STA thread via IGUIProxy | Must |

### Layout Operations

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-002-008 | Page SHALL display a list of all layouts in the current drawing (excluding "Model") | Must |
| FR-002-009 | Each layout entry SHALL show layout name and tab order | Must |
| FR-002-010 | User SHALL be able to select a layout to view its details | Must |
| FR-002-011 | Page SHALL show the currently active layout highlighted | Should |
| FR-002-012 | User SHALL be able to switch to a different layout | Should |

### Parameter Extraction

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-002-013 | Page SHALL display extracted PR Number from the active layout | Must |
| FR-002-014 | Page SHALL display Project Name, Job Working Plan Name from layout attributes | Must |
| FR-002-015 | Page SHALL show extracted block attributes (product_name, spec, category, etc.) | Must |
| FR-002-016 | Page SHALL display table data from valid layout tables in a DataGrid | Must |
| FR-002-017 | Table data SHALL show: position, product_no, width, height, len, thickness, qty, desc | Must |
| FR-002-018 | Page SHALL validate table structure (9 columns, HEADER_ID in col 7) | Must |
| FR-002-019 | Empty rows (both qty and product_no empty) SHALL be filtered out | Must |
| FR-002-020 | Text formatting codes SHALL be stripped via LM_UnFormat equivalent | Must |

### Write-Back Operations

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-002-021 | User SHALL be able to update block attributes in AutoCAD from the UI | Must |
| FR-002-022 | User SHALL be able to clear table IDs for the current layout | Must |
| FR-002-023 | User SHALL be able to clear table IDs for ALL layouts | Must |
| FR-002-024 | Write-back operations SHALL confirm success/failure to the user | Must |

### Drawing Operations (Advanced)

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-002-025 | Page SHOULD support creating new drawings with template and unit selection | Could |
| FR-002-026 | Page SHOULD support drawing primitives (line, circle, text) | Could |
| FR-002-027 | Page SHOULD support adding dimensions | Could |
| FR-002-028 | Page SHOULD support layer management (create, set current, list) | Could |
| FR-002-029 | Page SHOULD support entity scanning | Could |

## 5. UI Wireframe Description

```
+--------------------------------------------------------------+
|                    AutoCAD Integration                        |
+--------------------------------------------------------------+
|                                                               |
|  Connection Status                                            |
|  +----------------------------------------------------------+|
|  | * Connected to AutoCAD 2024                               ||
|  | Document: C:\Projects\Steel\Frame-001.dwg                 ||
|  | PR No: PR-2025-001 | Project: Steel Frame Phase 2        ||
|  | [Disconnect]  [Refresh Status]                            ||
|  +----------------------------------------------------------+|
|                                                               |
|  Layouts                    |  Layout Details                 |
|  +-----------------------+  |  +----------------------------+ |
|  | > S405-201            |  |  | Block Attributes:          | |
|  |   S405-202            |  |  | PR No: PR-2025-001        | |
|  |   S405-203 (active)   |  |  | Project: Steel Frame Ph2  | |
|  |   S405-204            |  |  | Job Plan: JP-001          | |
|  |                       |  |  | Material: H-BEAM          | |
|  |                       |  |  | Spec: SS400               | |
|  |                       |  |  | Color: RAL 7035           | |
|  +-----------------------+  |  +----------------------------+ |
|                                                               |
|  Table Data                                                   |
|  +----------------------------------------------------------+|
|  | Pos | Product   | W   | H   | L    | T  | Qty | Desc    ||
|  | A1  | H-200x200 | 200 | 200 | 6000 | 12 | 4   | Main bm||
|  | A2  | H-150x150 | 150 | 150 | 4000 | 10 | 8   | Sub bm ||
|  | A3  | PL-12     | 300 | 300 |      | 12 | 16  | Base pl||
|  +----------------------------------------------------------+|
|                                                               |
|  Actions                                                      |
|  [Extract Parameters] [Clear IDs (This Layout)] [Clear All]  |
|                                                               |
+--------------------------------------------------------------+
```

### Layout Details
- **Connection Panel**: Top section with status, document info, buttons
- **Layouts List**: Left panel ListView with layout names, current active highlighted
- **Layout Details**: Right panel showing extracted block attributes as key-value pairs
- **Table Data**: Full-width DataGrid showing extracted table rows
- **Actions Bar**: Bottom row with action buttons

## 6. Data Model

### ViewModel Properties

```csharp
public class AutoCADViewModel : ObservableObject
{
    // Connection
    public bool IsConnected { get; set; }
    public AutoCADStatus Status { get; set; }
    public string DocumentName { get; set; }
    public string DocumentPath { get; set; }

    // Project Context
    public string PRNumber { get; set; }
    public string ProjectName { get; set; }
    public string JobWorkingPlanName { get; set; }
    public int? ProjectId { get; set; }

    // Layouts
    public ObservableCollection<LayoutInfo> Layouts { get; set; }
    public LayoutInfo SelectedLayout { get; set; }

    // Layout Attributes
    public Dictionary<string, string> LayoutAttributes { get; set; }

    // Table Data
    public ObservableCollection<TableRowData> TableRows { get; set; }
    public string HeaderId { get; set; }

    // Commands
    public IAsyncRelayCommand ConnectCommand { get; }
    public IAsyncRelayCommand DisconnectCommand { get; }
    public IRelayCommand RefreshStatusCommand { get; }
    public IRelayCommand<LayoutInfo> SelectLayoutCommand { get; }
    public IAsyncRelayCommand ExtractParametersCommand { get; }
    public IAsyncRelayCommand ClearTableIdsCommand { get; }
    public IAsyncRelayCommand ClearAllTableIdsCommand { get; }
}

public class TableRowData
{
    public string Position { get; set; }
    public string ProductNo { get; set; }
    public string Width { get; set; }
    public string Height { get; set; }
    public string Length { get; set; }
    public string Thickness { get; set; }
    public string Quantity { get; set; }
    public string Description { get; set; }
    public string DetailId { get; set; }
}
```

## 7. API/Service Dependencies

| Service | Interface | Methods Used |
|---------|-----------|-------------|
| AutoCAD Service | `IAutoCADService` | `ConnectAsync()`, `DisconnectAsync()`, `GetStatusAsync()`, `GetLayouts()`, `GetLayoutsValues()`, `GetLayoutValues(name)`, `SetTableValue()`, `GetTables()`, `ClearSelection()` |
| GUI Proxy | `IGUIProxy` | `ExecuteInGuiAsync()` for all COM operations |
| Odoo Service | `IOdooService` | `GetProjectAsync()` for PR-to-Project lookup |
| Navigation | `INavigationService` | Return navigation |

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-002-001 | Connect button only enabled when not connected |
| VR-002-002 | All COM-dependent controls disabled when disconnected |
| VR-002-003 | Tables must have exactly 9 columns to be considered valid |
| VR-002-004 | Column 7 header must contain "HEADER_ID" for table to be valid |
| VR-002-005 | Rows where both qty and product_no are empty must be filtered |
| VR-002-006 | Clear operations require confirmation dialog |
| VR-002-007 | Extract Parameters requires both AutoCAD connected and project context |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| AutoCAD not running | "AutoCAD is not running. Please start AutoCAD and try again." | Show error, suggest start |
| COM connection failed | "Failed to connect to AutoCAD. Error: {details}" | Show error with details |
| No active document | "No drawing is open in AutoCAD. Please open a drawing." | Disable layout/table controls |
| ActiveDocument retry exhausted | "Could not access the active document after 5 attempts." | Show retry option |
| Invalid table structure | "Table in layout '{name}' has invalid structure (expected 9 columns)." | Skip table, log warning |
| PR not found in Odoo | "PR number '{pr_no}' not found in Odoo projects." | Show warning, continue |
| COM thread conflict | "AutoCAD operation timed out. The application may be busy." | Retry via GUI proxy |
| Attribute write failed | "Failed to update attribute '{tag}' in AutoCAD." | Show error, rollback |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/AutoCADPage.xaml` - WPF page
- `ViewModels/AutoCADViewModel.cs` - MVVM ViewModel
- `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` (exists, 396 lines)
- `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` (implementation)
- `OdooAutoCAD.Core/Threading/IGUIProxy.cs` (exists)

### Thread Safety (Critical) — Updated 2026-02-11

- ALL AutoCAD COM operations MUST go through `IGUIProxy.ExecuteInGuiAsync()`
- WPF uses `DispatcherTimer` (100ms interval) to call `IGUIProxy.ProcessRequests()`
- COM requires STA thread; MCP server runs on MTA thread — GUI proxy bridges this gap
- Python pattern: `gui_proxy.execute_in_gui("action", **kwargs)` maps to C#: `IGUIProxy.ExecuteInGuiAsync("action", params)`
- **Pure STA Mode**: All GUIProxy handlers execute COM operations directly on the STA thread (no `Task.Run`). Sync operations use `Task.FromResult`; `ConnectInternalAsync` uses `await Task.Delay` for retries.
- **OleMessageFilter**: Registered on STA thread at startup (`App.xaml.cs`) to handle `RPC_E_CALL_REJECTED` when AutoCAD is busy with WPF layout processing.
- **GUIProxy fast-path/slow-path**: `ProcessSingleRequest` checks `IsCompletedSuccessfully` for instant completion (sync handlers), or schedules `ContinueWith` on STA `SynchronizationContext` (async handlers like connect) to avoid blocking the STA thread.
- **IDispatch QI warmup**: `GetActiveObject()` calls `Marshal.GetIDispatchForObject()` immediately after obtaining the COM object to prevent deferred cross-process QueryInterface crashes.

### COM Connection Pattern (C#) — Updated 2026-02-11
```csharp
// ConnectAsync routes through GUIProxy → ConnectInternalAsync runs on STA
public async Task<bool> ConnectAsync()
{
    var response = await _guiProxy.ExecuteInGuiAsync("autocad_connect", null, timeout: 15000);
    return response.Success && response.Result is bool connected && connected;
}

// Handler executes directly on STA thread (no Task.Run):
_guiProxy.RegisterHandler("autocad_connect", async (parameters) =>
{
    return await ConnectInternalAsync();
    // 1. P/Invoke GetActiveObject("AutoCAD.Application") + IDispatch QI warmup
    // 2. Retry ActiveDocument up to 5 times with await Task.Delay(1000)
    // OleMessageFilter handles RPC_E_CALL_REJECTED automatically
});
```

### Text Formatting Cleanup
- Python's `LM_UnFormat()` removes AutoCAD MText formatting codes
- C# equivalent needed: regex to strip `\P`, `\C`, `\F`, `\H`, `\S`, brace formatting
- Used on all table cell values before display

### Table Column Mapping
```
Index 0: Position    -> "position"
Index 1: Product No  -> "product_no"
Index 2: Width       -> "width"
Index 3: Height      -> "height"
Index 4: Length      -> "len"
Index 5: Thickness   -> "thickness"
Index 6: Quantity    -> "qty"
Index 7: Description -> "desc"
Index 8: Detail ID   -> "detail_id" (hidden, used for Odoo sync)
```

### Key Differences from Python
- Python uses `pythoncom.CoInitialize()` — C# uses STA thread via DispatcherTimer
- Python buttons directly call COM — C# all goes through IGUIProxy (Pure STA Mode)
- Python shows parameter form in main content area — C# uses dedicated page with panels
- Python `GetActiveObject` — C# P/Invoke `oleaut32.dll!GetActiveObject` + IDispatch QI warmup (since `Marshal.GetActiveObject` removed in .NET Core)
- Python `Dispatch` — C# `Activator.CreateInstance(Type.GetTypeFromProgID(...))`
- Python has no COM message filter — C# registers `OleMessageFilter` for `RPC_E_CALL_REJECTED` retry
- Python uses `Thread.Sleep` for retry — C# uses `await Task.Delay` to keep UI responsive
