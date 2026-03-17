# COM/IPC 模式隔離架構

> **版本**: 1.0
> **日期**: 2026-03-17

## 問題

v6.0 新增 IPC 模式過程中，COM 模式出現多項回歸：
- `push_to_boq` 在 background thread 執行 COM 物件 → STA crash
- `get_doc_layouts()` COM 端回傳 COM 物件而非字串
- `get_block_attributes()` / `get_single_layout_values()` COM 端缺失
- `get_layouts_values()` 格式不一致 (COM `{"all":[...]}` vs IPC `[...]`)
- GUI Proxy 中多處 `.Name` 存取假設回傳 COM 物件

## 解決方案

### 1. 介面契約 (ABC)

`utility/autocad_backend_interface.py` 定義 `AutoCADBackendInterface(ABC)`：

```
AutoCADBackendInterface
├── connected_autocad() → bool
├── connect_autocad(main_body=None)
├── get_active_layout() → Optional[str]      # 字串，非 COM 物件
├── get_doc_layouts() → List[str]             # 字串列表
├── get_layouts_values() → {"all": [...]}     # 統一格式
├── get_single_layout_values(name) → dict
├── get_block_attributes() → dict
├── set_block_attributes(attrs, layout_name) → bool
├── set_layouts_tables_id(boq_list)
├── get_layouts_header_id_to_pr() → {"all": [...]}
├── clear_table_id(layout=None)
├── clear_all_tables_id()
├── draw_line / draw_circle / set_layer / list_layers / scan_elements
```

`UtilAutoCAD` (COM) 和 `UtilAutoCADIPC` (IPC) 都繼承此 ABC。
缺方法 → Python 拋 `TypeError`，不需 `hasattr()` 防護。

### 2. 回傳格式統一

| 方法 | 回傳格式 |
|------|----------|
| `get_active_layout()` | `str` 或 `None` |
| `get_doc_layouts()` | `List[str]` |
| `get_layouts_values()` | `{"all": [layout_dict, ...]}` |
| `get_layouts_header_id_to_pr()` | `{"all": [header_id, ...]}` |

COM 內部使用 `get_layout_from_name(name)` 取得 COM Layout 物件。

### 3. 模式感知線程策略

```
_run_with_progress()
├── COM → _run_with_progress_main_thread()
│   └── 主線程同步執行 + update_idletasks()
└── IPC → _run_with_progress_background()
    └── 背景 Thread + ProgressDialog + Queue
```

COM 模式在主線程執行，確保 STA 安全。
IPC 模式使用背景線程，避免 UI 凍結（IPC 操作可能秒級以上）。

### 4. 業務邏輯解耦

`util_push_to_boq.py` 不再檢查 `mode == 'ipc'`：
- 統一嘗試 per-layout → fallback 批次
- 兩端 `get_layouts_values()` 都回傳 `{"all": [...]}`

## 修改清單

| 檔案 | 變更 |
|------|------|
| `utility/autocad_backend_interface.py` | 新增 ABC |
| `utility/util_autocad.py` | 繼承 ABC，新增 `get_block_attributes` / `get_single_layout_values`，修正 `get_doc_layouts` 回傳字串 |
| `utility/util_autocad_ipc.py` | 繼承 ABC，修正 `get_layouts_values` 回傳 `{"all": [...]}` |
| `utility/util_autocad_dispatcher.py` | 移除 `get_single_layout_values` 的 `hasattr` 防護 |
| `utility/util_push_to_boq.py` | 移除 `is_ipc` 模式檢查，統一流程 |
| `forms/form_main_modern.py` | 拆分 `_run_with_progress` 為 COM/IPC 兩種策略 |
| `utility/util_gui_proxy.py` | 移除 `.Name` 存取（layout 已是字串）|
