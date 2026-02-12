# FR-009: Odoo Parameter Configuration (從 Odoo 獲取參數配置)

> **Document Version**: 1.0
> **Last Updated**: 2026-02-12
> **Status**: Not Started
> **Priority**: P2

## 1. Overview

The Odoo Parameter Configuration feature enables users to fetch material, processing, and color options from Odoo via Swagger API, present them in a form with searchable dropdowns, and write selected values to AutoCAD attribute blocks. This bridges the Odoo product/setup data with AutoCAD drawing parameters, allowing engineers to configure drawing attributes directly from the application without manual data entry.

This feature corresponds to the Python `forms/form_autocad_param.py` "從 Odoo 獲取參數" workflow and is self-contained — it depends only on Sprint 3 infrastructure (AutoCAD connection, Odoo connection, Swagger API patterns) which is already complete.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-009-01 | CAD Engineer | Fetch material/setup/color options from Odoo via API | I have up-to-date dropdown data for my drawing parameters |
| US-009-02 | CAD Engineer | See a parameter configuration form with searchable dropdowns | I can select the correct material, processing, and color values |
| US-009-03 | CAD Engineer | Write selected parameter values to AutoCAD attribute blocks | The drawing attributes are updated with the correct Odoo data |

## 3. Python Reference

### Source Files
- `forms/form_autocad_param.py` (~234 lines) — 3-section form (材料/加工/顏色) with PopupSelector widgets
- `utility/util_odoo.py` (~600 lines) — `get_product_v2`, `get_setup_v2`, `get_color_v2` via Swagger API
- `utility/util_autocad.py` (~1,301 lines) — `set_attribute_value(block, tag, val)`, `get_attribute_block(layout)`

### Key Functions
- `get_product_v2` — Fetches products via Swagger `callMethodForJobWorkingPlanBoqModel` with filter `categ_id child_of 27, active=True`
- `get_setup_v2(setup_name)` — Fetches setup values (spec, product_catelog, operation_flow, surface_treatment) with filter `setup_name = {name}`
- `get_color_v2(project_id)` — Fetches project-specific colors with filter `job_project_id = {project_id}`
- `set_attribute_value(block, tag, val)` — Sets a single attribute on an AutoCAD block reference
- `get_attribute_block(layout)` — Finds the attribute block in a layout (block reference with attributes)

### Form Layout (Python)
```
+--------------------------------------------------------------+
| 從 Odoo 獲取參數                                              |
+--------------------------------------------------------------+
| 材料配置                                                      |
|   材料:    [material_entry    ▼]  單位: [unit (RO)]           |
|   材質:    [spec_entry        ▼]                               |
|   材料分類: [category_entry   ▼]                               |
+--------------------------------------------------------------+
| 加工配置                                                      |
|   加工流程: [process_entry    ▼]                               |
|   表面處理: [surface_entry    ▼]                               |
+--------------------------------------------------------------+
| 顏色配置                                                      |
|   顏色:    [color_entry       ▼]  色號: [color_no (RO)]       |
+--------------------------------------------------------------+
|          [確定提交]              [取消]                         |
+--------------------------------------------------------------+
```

### Form Fields → AutoCAD Attribute Tags

| Section | Field | Odoo API | Method Name | AutoCAD Tag |
|---------|-------|----------|-------------|-------------|
| 材料配置 | 材料 + 單位(RO) | get_product_v2 | `get_product_v2` | `product_name` |
| 材料配置 | 材質 | get_setup_v2('spec') | `get_setup_v2` | `spec` |
| 材料配置 | 材料分類 | get_setup_v2('product_catelog') | `get_setup_v2` | `product_catelog` |
| 加工配置 | 加工流程 | get_setup_v2('operation_flow') | `get_setup_v2` | `operation_flow` |
| 加工配置 | 表面處理 | get_setup_v2('surface_treatment') | `get_setup_v2` | `surface_treatment` |
| 顏色配置 | 顏色 + 色號(RO) | get_color_v2(project_id) | `get_color_v2` | `color_name` / `color_no` |

### Submit Flow (Python)
1. Validate all 7 fields are non-empty
2. Get attribute block from active layout: `get_attribute_block(layout)`
3. Iterate all layouts, call `set_attribute_value(block, tag, val)` for each tag:
   - `product_name`, `spec`, `product_catelog`, `operation_flow`, `surface_treatment`, `color_name`, `color_no`
4. Log success or error message

## 4. Functional Requirements

### API Data Fetching

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-009-001 | Page SHALL fetch product options from Odoo via `get_product_v2` Swagger endpoint | Must |
| FR-009-002 | Page SHALL fetch setup values (spec, product_catelog, operation_flow, surface_treatment) via `get_setup_v2` Swagger endpoint | Must |
| FR-009-003 | Page SHALL fetch project-specific color options via `get_color_v2` Swagger endpoint | Must |
| FR-009-004 | All API calls SHALL use BasicAuth + Swagger PATCH pattern established in Sprint 5/6 | Must |
| FR-009-005 | API data SHALL be fetched in parallel on page load to minimize wait time | Should |
| FR-009-006 | API failures SHALL show user-friendly error messages without crashing | Must |

### Form UI

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-009-007 | Page SHALL display a 3-section form: 材料配置, 加工配置, 顏色配置 | Must |
| FR-009-008 | Material field SHALL be a searchable ComboBox populated from product API data | Must |
| FR-009-009 | Unit field SHALL be read-only, auto-filled when material is selected | Must |
| FR-009-010 | Spec, Category, Operation Flow, Surface Treatment fields SHALL be searchable ComboBoxes | Must |
| FR-009-011 | Color field SHALL be a searchable ComboBox; Color No SHALL be read-only auto-filled | Must |
| FR-009-012 | Submit button SHALL be enabled only when all 7 required fields are filled | Must |
| FR-009-013 | Cancel button SHALL clear all fields and reset the form | Should |
| FR-009-014 | Page SHALL be accessible via sidebar navigation button | Must |

### Write to AutoCAD

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-009-015 | Submit SHALL write all 7 attribute values to the attribute block in each layout | Must |
| FR-009-016 | Write operation SHALL execute on the GUI/STA thread via IGUIProxy | Must |
| FR-009-017 | Write operation SHALL require AutoCAD to be connected | Must |
| FR-009-018 | Write operation SHALL display success/failure feedback to the user | Must |
| FR-009-019 | Write operation SHALL find the attribute block via `get_attribute_block` equivalent | Must |

## 5. UI Wireframe Description

```
+--------------------------------------------------------------+
|                    參數配置                                    |
+--------------------------------------------------------------+
|                                                               |
|  材料配置                                                     |
|  +----------------------------------------------------------+|
|  | 材料:     [Searchable ComboBox          ▼]               ||
|  | 單位:     [Read-only TextBox            ]                 ||
|  | 材質:     [Searchable ComboBox          ▼]               ||
|  | 材料分類:  [Searchable ComboBox          ▼]               ||
|  +----------------------------------------------------------+|
|                                                               |
|  加工配置                                                     |
|  +----------------------------------------------------------+|
|  | 加工流程:  [Searchable ComboBox          ▼]               ||
|  | 表面處理:  [Searchable ComboBox          ▼]               ||
|  +----------------------------------------------------------+|
|                                                               |
|  顏色配置                                                     |
|  +----------------------------------------------------------+|
|  | 顏色:     [Searchable ComboBox          ▼]               ||
|  | 色號:     [Read-only TextBox            ]                 ||
|  +----------------------------------------------------------+|
|                                                               |
|  [確定提交]                                    [取消]          |
|                                                               |
|  Status: Ready / Loading options... / Writing to AutoCAD...  |
|                                                               |
+--------------------------------------------------------------+
```

### Layout Details
- **Material Section**: GroupBox with 4 fields (material, unit, spec, category)
- **Processing Section**: GroupBox with 2 fields (operation_flow, surface_treatment)
- **Color Section**: GroupBox with 2 fields (color, color_no)
- **Action Bar**: Submit and Cancel buttons
- **Status Area**: Status message for loading/writing feedback

## 6. Data Model

### ViewModel Properties

```csharp
public class ParameterConfigViewModel : ObservableObject
{
    // Material Section
    public ObservableCollection<OdooProduct> Products { get; set; }
    public OdooProduct? SelectedProduct { get; set; }
    public string Unit { get; set; } // Read-only, auto-filled

    public ObservableCollection<OdooSetupValue> Specs { get; set; }
    public OdooSetupValue? SelectedSpec { get; set; }

    public ObservableCollection<OdooSetupValue> Categories { get; set; }
    public OdooSetupValue? SelectedCategory { get; set; }

    // Processing Section
    public ObservableCollection<OdooSetupValue> OperationFlows { get; set; }
    public OdooSetupValue? SelectedOperationFlow { get; set; }

    public ObservableCollection<OdooSetupValue> SurfaceTreatments { get; set; }
    public OdooSetupValue? SelectedSurfaceTreatment { get; set; }

    // Color Section
    public ObservableCollection<OdooColor> Colors { get; set; }
    public OdooColor? SelectedColor { get; set; }
    public string ColorNo { get; set; } // Read-only, auto-filled

    // State
    public bool IsLoading { get; set; }
    public bool IsSubmitting { get; set; }
    public string StatusMessage { get; set; }
    public bool AllFieldsFilled { get; } // Computed

    // Commands
    public IAsyncRelayCommand LoadOptionsCommand { get; }
    public IAsyncRelayCommand SubmitCommand { get; }
    public IRelayCommand CancelCommand { get; }
}
```

### New Record Types

```csharp
public record OdooSetupValue(string Value, string SetupName);
public record OdooColor(string Name, string ColorNo, int ProjectId);
```

## 7. API/Service Dependencies

| Service | Interface | Methods Used |
|---------|-----------|-------------|
| Odoo Service | `IOdooService` | `GetProductsViaApiAsync()` (existing), `GetSetupViaApiAsync()` (new), `GetColorsViaApiAsync()` (new) |
| AutoCAD Service | `IAutoCADService` | `autocad_get_attribute_block` (new handler), `autocad_set_attribute_values` (new handler) |
| GUI Proxy | `IGUIProxy` | `ExecuteInGuiAsync()` for all COM operations |
| Settings Service | `ISettingsService` | Swagger URL, database, user token retrieval |
| Navigation | `INavigationService` | Sidebar navigation to/from page |

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-009-001 | Submit button only enabled when all 7 fields are filled (AllFieldsFilled) |
| VR-009-002 | Submit requires AutoCAD to be connected |
| VR-009-003 | Color dropdown requires valid project_id (from AutoCAD layout attributes) |
| VR-009-004 | Read-only fields (Unit, ColorNo) cannot be edited directly |
| VR-009-005 | API calls require valid Odoo connection settings (Swagger URL, database, token) |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| Odoo API unreachable | "無法連接 Odoo 伺服器，請檢查連線設定。" | Show error, disable dropdowns |
| Product fetch failed | "無法載入材料清單: {details}" | Show error, retry option |
| Setup fetch failed | "無法載入設定值 ({type}): {details}" | Show error for specific type |
| Color fetch failed | "無法載入顏色清單: {details}" | Show error, color section disabled |
| No project_id | "無法取得專案 ID，請確認 AutoCAD 圖檔已設定 PR 編號。" | Show warning, disable color section |
| AutoCAD not connected | "AutoCAD 未連線，請先連接 AutoCAD。" | Disable Submit button |
| Attribute block not found | "在佈局 '{name}' 中未找到屬性區塊。" | Show error, skip layout |
| Write failed | "更新 AutoCAD 屬性時發生錯誤: {details}" | Show error, log details |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/ParameterConfigPage.xaml` — WPF page (new)
- `Views/Pages/ParameterConfigPage.xaml.cs` — Code-behind with DI (new)
- `ViewModels/ParameterConfigViewModel.cs` — MVVM ViewModel (new)
- `OdooAutoCAD.Core/Odoo/IOdooService.cs` — Add 2 new methods
- `OdooAutoCAD.Core/Odoo/OdooService.cs` — Implement new methods
- `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` — Add 2 new GUIProxy handlers
- `Views/MainWindow.xaml` — Add sidebar navigation button
- `App.xaml.cs` — Register ParameterConfigViewModel in DI

### C# Reusable Code
- `GetProductsViaApiAsync()` — Already exists for materials, can be reused directly
- `ParseSwaggerUrl()` / `ResolveSwaggerEndpoint()` — Swagger endpoint resolution (in OdooConnectionViewModel)
- PATCH + BasicAuth pattern — Established in `ImportToBOQViaApiAsync`, `ConvertBOQToPRViaApiAsync`
- `GetAttributeValues()` in AutoCADService — Reads block attributes (reference for write)
- GUIProxy handler pattern — 15+ handlers already registered
- Navigation + DI + page pattern — Fully established

### What's New (to implement)
1. `IOdooService`: `GetSetupViaApiAsync()`, `GetColorsViaApiAsync()`, `OdooSetupValue`, `OdooColor` records
2. `OdooService`: Implementations following `GetProductsViaApiAsync` pattern
3. `AutoCADService`: `autocad_set_attribute_values` + `autocad_get_attribute_block` handlers
4. `ParameterConfigViewModel` — New ViewModel with dropdown collections and submit logic
5. `ParameterConfigPage.xaml` — New page with 3 GroupBox sections
6. Sidebar button + DI registration
7. Tests (~24 new tests)

### Thread Safety (Critical)
- All AutoCAD attribute write operations MUST go through `IGUIProxy.ExecuteInGuiAsync()`
- The `autocad_set_attribute_values` handler iterates all layouts and sets attributes on the STA thread
- No `Task.Run` — follows the Pure STA Mode pattern established in Sprint 3

### Key Differences from Python
- Python uses `PopupSelector` (custom Tkinter popup) — C# uses `ComboBox` with `IsEditable=True` for search
- Python fetches API data lazily on field focus — C# loads all options in parallel on page load
- Python calls `set_attribute_value` per layout per tag — C# batches all tags into a single handler call
- Python uses `get_attribute_block(layout)` per layout — C# handler iterates layouts internally
