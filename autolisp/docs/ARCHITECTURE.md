# AutoLISP + IPC/MCP 架構設計

> 版本: 4.0 (雙模式架構 — DCL 獨立模式已移除)
> 日期: 2026-03-10
> 架構: 雙模式 — COM / IPC+MCP (autocad-mcp)
> 實作計畫: 見 [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md)

---

## 1. 架構概覽

本專案支援兩條操作路徑，以 `puran-water/autocad-mcp` 作為 git submodule 提供通用 AutoCAD 工具基礎：

1. **COM Mode** — Python GUI + COM 直連 Full AutoCAD（現有，不改）
2. **IPC Mode** — Python GUI + File IPC 驅動 AutoCAD LT 2024+（基於 autocad-mcp submodule）
3. **MCP Mode** — AI 助手透過 MCP Server + File IPC 驅動 AutoCAD（基於 autocad-mcp submodule）

```
┌──────────────────────┐  ┌──────────────────────┐  ┌──────────────────────────┐
│  Python GUI          │  │  Python GUI          │  │  AI Assistant            │
│  (COM Mode)          │  │  (IPC Mode)          │  │  (MCP Mode)              │
│  Settings: ●COM ○IPC │  │  Settings: ○COM ●IPC │  │  Claude / Gemini CLI     │
│                      │  │                      │  │                          │
│  util_autocad.py     │  │  util_autocad_ipc.py │  │  MCP Client (stdio/SSE)  │
│  (COM 直連)          │  │  (wraps autocad-mcp) │  │                          │
└──────────┬───────────┘  └──────────┬───────────┘  └──────────┬───────────────┘
           │                         │                          │
      COM Automation            File IPC                  MCP Protocol
           │               (autocad-mcp lib)                    │
           ▼                         │               ┌──────────▼──────────────┐
┌──────────────────┐                 │               │  MCP Server              │
│  Full AutoCAD    │                 │               │  (wraps autocad-mcp      │
│  (COM objects)   │                 │               │   + Odoo tools)          │
└──────────────────┘                 │               │                          │
                                     │               │  8 通用 tools (autocad)  │
                                     │               │  5 Odoo tools (我們的)    │
                                     │               └──────────┬──────────────┘
                                     │                          │
                                     ▼                          ▼
                              ┌──────────────────────────────────────┐
                              │  AutoCAD LT 2024+ (or Full)          │
                              │                                      │
                              │  autocad-mcp/lisp-code/              │
                              │    mcp_dispatch.lsp  (通用 dispatcher)│
                              │                                      │
                              │  autolisp/lisp/                      │
                              │    070_ob_mcp_dispatch.lsp (Odoo 擴展)│
                              │    060_table_util.lsp (TABLE 操作)    │
                              │    050_block_util.lsp (Block 操作)    │
                              └──────────────────────────────────────┘
```

### Python GUI 內部架構（COM/IPC 切換）

```
┌─────────────────────────────────────────────────────────────┐
│  odoo.py  (進入點)                                          │
│  --autocad-mode com|ipc    (預設 com)                        │
│  --mcp-autocad             (啟動 autocad MCP Server)         │
└──────────────────────┬──────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────┐
│  forms/form_main_modern.py  (GUI 主視窗)                     │
│                                                              │
│  init_utilities():                                           │
│  ┌─────────────────────────────────────────────────┐         │
│  │  UtilAutoCADDispatcher (mode=com|ipc)            │         │
│  │  ┌──────────────┐  ┌───────────────────────┐    │         │
│  │  │ UtilAutoCAD  │  │ UtilAutoCADIPC        │    │         │
│  │  │ (COM 後端)   │  │ (IPC 後端)            │    │         │
│  │  │ pywin32 COM  │  │ wraps autocad-mcp lib │    │         │
│  │  └──────┬───────┘  └───────────┬───────────┘    │         │
│  │         │ (COM Mode)           │ (IPC Mode)     │         │
│  └─────────┼──────────────────────┼────────────────┘         │
│            │                      │                           │
│  ┌─────────▼──────────┐          │                           │
│  │ GUIProxy           │          │ (IPC 不需 GUI proxy,      │
│  │ (COM 線程安全)     │          │  無 COM 線程問題)          │
│  │ 100ms timer 輪詢   │          │                           │
│  └────────────────────┘          │                           │
│                                   │                           │
│  UtilPushToBoq ──→ dispatcher     │                           │
│  UtilTransferBoqToPr ──→ dispatcher                           │
│  MCPManager ──→ dispatcher + mcp_server_autocad.py            │
└──────────────────────┬───────────┬───────────────────────────┘
                       │           │
                  COM Automation   File IPC
                       │           │
                       ▼           ▼
              Full AutoCAD    AutoCAD LT 2024+
```

**GUI → MCP 層替換:**

| 舊 (v5.0 SSE 模式) | 新 (Phase 6) |
|---|---|
| `MCPSSEManager` (`util_mcp_sse_manager.py`) | `MCPManager` (`util_mcp_manager.py`) — 全新重寫 |
| `mcp_server_fastmcp.py` (7 tools) | `mcp_server_autocad.py` (13 tools) |
| `util_mcp_sse_server.py` (SSE transport) | `mcp_server_autocad.py` 內建 stdio/streamable-http |
| `set_shared_autocad_util()` | 接收 `dispatcher` + `odoo_util` |

---

## 2. 目錄結構（已實作 + Phase 6 新增）

```
autolisp/
├── README.md                     # 安裝與使用說明
├── docs/                         # 文件
│   ├── ARCHITECTURE.md           # 本文件（架構 + 模組規格）
│   ├── IMPLEMENTATION_PLAN.md    # 實作計畫 + 進度追蹤
│   └── AUTOCAD_LT_RESEARCH.md   # LT 支援研究報告
│
├── lisp/                         # AutoLISP 原始碼（LT 相容）
│   ├── 080_main.lsp              # 進入點，載入所有模組，定義使用者指令
│   ├── 030_config.lsp            # 路徑常數、YAML 設定讀取
│   ├── 010_json_util.lsp         # JSON 解析/序列化（diegomcas 版本）
│   ├── 020_file_util.lsp         # 檔案讀寫 + 唯一檔名 + 刪除
│   ├── 060_table_util.lsp        # TABLE 實體讀寫（遍歷 Layout, 讀取/回寫 ID）
│   ├── 050_block_util.lsp        # Block 屬性讀寫（get/set attribute）
│   ├── 040_strip_mtext.lsp       # MText 格式清除（RegExp + LT 純字串 fallback）
│   └── 070_ob_mcp_dispatch.lsp   # Odoo IPC 擴展 dispatcher
│
├── config/                       # 設定檔
│   ├── server_prod.yaml.example  # Odoo 伺服器設定範例
│   └── token.yaml.example        # 使用者 token 設定範例
│
└── legacy/                       # 舊版檔案（參考用，不再直接使用）
    ├── call_python.lsp           # 舊版 Python COM 呼叫
    ├── transfer_to_odoo.lsp      # 舊版 Odoo 整合（COM 方式）
    ├── contract_product.lsp      # 舊版參數選擇
    ├── contract_product.dcl      # 舊版 DCL
    ├── param_form.dcl            # 舊版 DCL
    ├── json_util.lsp             # JSON 工具（已移入 lisp/）
    ├── read_csv.lsp              # CSV 工具（已棄用，改由 API 取資料）
    └── StripMtext v5-0b.lsp      # MText 工具（已移入 lisp/ 並加 LT fallback）

libs/                             # Git submodules
└── autocad-mcp/                  # puran-water/autocad-mcp
     ├── src/autocad_mcp/         → Python: PostMessageW, File IPC, ezdxf backend
     ├── lisp-code/               → mcp_dispatch.lsp (通用 AutoCAD dispatcher)
     └── ...

# Phase 6 Python 端新增檔案（在專案根目錄）:
utility/util_autocad_ipc.py        # 包裝 autocad-mcp File IPC client
utility/util_autocad_dispatcher.py # COM/IPC 模式切換（統一介面）
mcp_server_autocad.py              # MCP Server（包裝 autocad-mcp + Odoo tools）
utility/util_mcp_manager.py        # 全新 MCP 管理器（取代 util_mcp_sse_manager.py）

# Phase 6 需修改的現有檔案:
forms/form_main_modern.py          # 改用 UtilAutoCADDispatcher，條件化 GUI proxy，改用 MCPManager
utility/util_gui_proxy.py          # COM 模式才啟用，IPC 模式 bypass
odoo.py                            # 新增 --autocad-mode / --mcp-autocad 參數

# Phase 6 刪除（SSE 舊模式）:
utility/util_mcp_sse_manager.py    # 刪除，由 util_mcp_manager.py 取代
utility/util_mcp_sse_server.py     # 刪除，由 mcp_server_autocad.py 內建 transport 取代
mcp_server_fastmcp.py              # 刪除，由 mcp_server_autocad.py 取代
```

---

## 3. 模組規格

### 3.1 lisp/030_config.lsp — 設定管理

**全域變數:**

| 變數 | 類型 | 說明 |
|------|------|------|
| `*ob:root*` | string | autolisp 根目錄路徑 |
| `*ob:lisp-dir*` | string | lisp/ 目錄路徑 |
| `*ob:config-dir*` | string | config/ 目錄路徑 |
| `*ob:server-yaml*` | string/nil | server_prod.yaml 路徑（優先使用） |
| `*ob:token-yaml*` | string/nil | token.yaml 路徑（優先使用） |
| `*ob:temp-dir*` | string | JSON 暫存目錄（%TEMP%/odoo_bridge/） |
| `*ob:odoo-connected*` | T/nil | Odoo 連線狀態 |
| `*ob:odoo-url*` | string | Odoo 伺服器 URL |
| `*ob:odoo-db*` | string | Odoo 資料庫名稱 |
| `*ob:odoo-user*` | string | 使用者名稱 |
| `*ob:odoo-pass*` | string | 密碼/API Key |

**設定檔優先順序:** YAML（server_prod.yaml + token.yaml）→ 自動偵測

**函數:**

| 函數 | 說明 |
|------|------|
| `(config:init)` | 初始化所有路徑，搜尋 YAML 設定 |

---

### 3.2 lisp/010_json_util.lsp — JSON 解析/序列化

基於 diegomcas 版本，修改 `value_to_string` 加入 `json:escape` 呼叫。

| 函數 | 說明 |
|------|------|
| `(dmc:json:json_to_list json quote_array)` | JSON 字串 → assoc list |
| `(dmc:json:list_to_json lst)` | assoc list → JSON 字串 |
| `(dmc:json:str_replace str patt repl_to)` | 字串全域替換 |

---

### 3.3 lisp/020_file_util.lsp — 檔案 I/O

| 函數 | 說明 |
|------|------|
| `(file:write-string filepath content)` | 寫字串到檔案 |
| `(file:write-lines filepath lines)` | 寫字串 list 到檔案 |
| `(file:read-string filepath)` | 讀檔案為單一字串 |
| `(file:read-lines filepath)` | 讀檔案為字串 list |
| `(file:unique-name dir prefix ext)` | 產生唯一檔名（含時間戳） |
| `(file:delete filepath)` | 刪除檔案 |
| `(file:ensure-dir dir)` | 確保目錄存在 |
| `(file:exists-p filepath)` | 檢查檔案是否存在 |
| `(json:escape str)` | JSON 序列化前跳脫特殊字元（`\` `"` LF CR TAB） |
| `(json:unescape str)` | JSON 解析後還原跳脫字元 |

---

### 3.4 lisp/060_table_util.lsp — TABLE 實體操作

**TABLE 結構（9 欄）:**

| Row | Col 0 | Col 1 | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 |
|-----|-------|-------|-------|-------|-------|-------|-------|-------|-------|
| 0 (title) | | | | | | | | HEADER_ID | header_id值 |
| 1 (header) | 位置 | 料號 | 寬 | 高 | 長 | 厚 | 數量 | 說明 | detail_id |
| 2+ (data) | position | product_no | width | height | length | thickness | qty | desc | detail_id |

**Cell 讀寫:**

| 函數 | 說明 |
|------|------|
| `(table:get-cell-text ename row col)` | 讀取 cell 文字（含 MText unformat） |
| `(table:set-cell-text ename row col value)` | 寫入 cell 文字 |
| `(table:get-rows ename)` | 取得列數 |
| `(table:get-columns ename)` | 取得欄數 |
| `(table:legal-p ename)` | 檢查是否為合法 TABLE（9 欄 + HEADER_ID） |

**資料收集:**

| 函數 | 說明 |
|------|------|
| `(table:get-layout-tables)` | 取得目前 Layout 所有合法 TABLE enames |
| `(table:get-detail-rows ename)` | 提取資料列（qty > 0） |
| `(table:get-header-id ename)` | 取得 header_id (0,8) |
| `(table:get-current-layout-data)` | 收集目前 Layout 的 TABLE 資料 |
| `(table:get-all-layouts-data)` | 收集所有 Layout 的 TABLE + Block 資料 |
| `(table:get-all-header-ids)` | 收集所有 Layout 的 header_id（用於 PR） |

**ID 回寫:**

| 函數 | 說明 |
|------|------|
| `(table:write-header-id ename id)` | 寫 header_id 到 (0,8) |
| `(table:write-detail-id ename product-no id)` | 寫 detail_id 到匹配的列 |
| `(table:update-ids-from-response response)` | 從回應批量回寫 ID |

**清除:**

| 函數 | 說明 |
|------|------|
| `(table:clear-ids ename)` | 清除單一 TABLE 的所有 ID |
| `(table:clear-current-layout-ids)` | 清除目前 Layout |
| `(table:clear-all-layouts-ids)` | 清除所有 Layout |

---

### 3.5 lisp/050_block_util.lsp — Block 屬性操作

**7 個屬性 Tag:**
`product_name`, `spec`, `product_catelog`, `operation_flow`, `surface_treatment`, `color_name`, `color_no`

**Header Tag:**
`job_working_plan_name`, `project_name`

| 函數 | 說明 |
|------|------|
| `(block:get-attribute obj tag)` | 讀取屬性值 |
| `(block:set-attribute obj tag value)` | 寫入屬性值 |
| `(block:get-all-attributes obj)` | 取得所有屬性 `((tag . value) ...)` |
| `(block:set-attributes obj attr-list)` | 批量寫入屬性 |
| `(block:find-attribute-block)` | 在目前 Layout 找到屬性 Block |
| `(block:find-attribute-block-in-layout layout)` | 在指定 Layout 找到屬性 Block |
| `(block:get-header-attrs block layout-name)` | 取得 header 屬性（用於 BOQ 匯出） |
| `(block:set-attributes-all-layouts attr-list)` | 寫入屬性到所有 Layout |

---

### 3.6 lisp/040_strip_mtext.lsp — MText 格式清除

| 函數 | 說明 |
|------|------|
| `(strip:unformat str)` | 統一入口：自動選擇 RegExp 或 fallback |
| `(LM:UnFormat str mtx)` | Lee Mac 版（VBScript.RegExp，Full AutoCAD） |
| `(strip:unformat-simple str)` | LT fallback（純 `vl-string-subst` 鏈式替換） |

**LT fallback 處理的格式碼:**
`\P`(段落), `\f`/`\F`(字型), `\C`/`\c`(顏色), `\H`(高度), `\W`(寬度), `\T`(追蹤), `\Q`(傾斜), `\A`(對齊), `\L`/`\l`(底線), `\O`/`\o`(刪除線), `\~`(不分行空格), `{}`(群組)

---

### 3.7 lisp/080_main.lsp — 進入點

**載入順序:**
1. 010_json_util.lsp
2. 020_file_util.lsp
3. 030_config.lsp
4. 040_strip_mtext.lsp
5. 050_block_util.lsp
6. 060_table_util.lsp
7. 070_ob_mcp_dispatch.lsp

**使用者指令（1 個）:**

| 指令 | 函數 | 說明 |
|------|------|------|
| `OB:MCP-DISPATCH` | `mcp:dispatch` | IPC/MCP dispatcher（命令列背景執行） |

---

### 3.8 lisp/070_ob_mcp_dispatch.lsp — Odoo IPC 擴展

此模組載入 autocad-mcp 的 `mcp_dispatch.lsp` 作為基礎 dispatcher，並擴展 dispatch table 加入 Odoo 專用 actions。

**設計原則:**
- autocad-mcp 的 `mcp_dispatch.lsp` 提供通用 dispatch 框架（讀取 JSON command → 執行 → 寫回 JSON result）
- `070_ob_mcp_dispatch.lsp` 在載入後追加 Odoo 擴展 actions 到 dispatch table
- Odoo actions 呼叫現有 `060_table_util.lsp` / `050_block_util.lsp` 執行 AutoCAD 操作

**擴展的 Odoo actions:**

| Action | 說明 | 呼叫的現有函數 |
|--------|------|---------------|
| `odoo_extract_tables` | 收集所有 Layout TABLE + Block 資料 | `table:get-all-layouts-data` |
| `odoo_get_header_ids` | 收集所有 Layout 的 header_id | `table:get-all-header-ids` |
| `odoo_write_ids` | 回寫 header_id + detail_id 到 TABLE | `table:update-ids-from-response` |
| `odoo_get_block_attrs` | 讀取屬性 Block 的所有屬性 | `block:find-attribute-block` + `block:get-all-attributes` |
| `odoo_set_block_attrs` | 寫入屬性到所有 Layout 的 Block | `block:set-attributes-all-layouts` |
| `odoo_clear_ids` | 清除 TABLE ID | `table:clear-all-layouts-ids` |

**載入流程:**
```lisp
;; 1. 載入 autocad-mcp 的通用 dispatcher
(load (strcat *ob:root* "/../libs/autocad-mcp/lisp-code/mcp_dispatch.lsp"))

;; 2. 擴展 dispatch table，加入 Odoo actions
(mcp:register-action "odoo_extract_tables" 'ob:action-extract-tables)
(mcp:register-action "odoo_get_header_ids" 'ob:action-get-header-ids)
;; ... 其他 Odoo actions
```

---

## 4. 安全考量

- **認證資訊** 存放在 YAML 設定檔（server_prod.yaml + token.yaml），不透過 JSON 檔案傳遞
- **設定檔搜尋路徑**: config/ → C:/odoo/config/ → autolisp root/
- **YAML 設定檔** 含 token，加入 `.gitignore`，提供 `.example` 範例
- **Token 明碼存放**（與現有 Python 主應用一致，共用同一組 YAML 設定）

---

## 5. 開發階段進度

### Phase 1-5: DCL 獨立模式 — DONE (已移除)

Phase 1-5 實作了 DCL 對話框 + Python Bridge.exe 的獨立模式架構。
此模式已被 Phase 6 的 IPC/MCP 模式取代，相關程式碼已刪除。
舊版檔案保留在 `legacy/` 目錄供參考。

### Phase 6: autocad-mcp Submodule + IPC/MCP Mode — TODO
- [ ] `libs/autocad-mcp/` — git submodule (puran-water/autocad-mcp)
- [ ] `lisp/070_ob_mcp_dispatch.lsp` — Odoo 擴展 dispatcher
- [ ] `utility/util_autocad_ipc.py` — Python File IPC client
- [ ] `utility/util_autocad_dispatcher.py` — COM/IPC 模式切換
- [ ] `mcp_server_autocad.py` — MCP Server（autocad-mcp + Odoo tools）
- [ ] `lisp/080_main.lsp` — 新增 OB:MCP-DISPATCH 指令

### 待辦（Future）
- [ ] AutoCAD 內端對端測試
- [ ] 安裝腳本（自動設定搜尋路徑）
- [ ] 使用者操作手冊

---

## 6. autocad-mcp Submodule 整合

### 6.1 Submodule 路徑與版本管理

```bash
# 新增 submodule
git submodule add https://github.com/puran-water/autocad-mcp.git libs/autocad-mcp

# 初始化（clone 後）
git submodule update --init --recursive

# 更新到最新版本
cd libs/autocad-mcp && git pull origin main && cd ../..
git add libs/autocad-mcp && git commit -m "chore: Update autocad-mcp submodule"
```

Git submodule：
- `libs/autocad-mcp/` — Python AutoCAD IPC + MCP tools（autolisp/Python 專案使用）

### 6.2 autocad-mcp 提供的 8 個通用 Tools

| Tool | 功能 | 說明 |
|------|------|------|
| `drawing` | 開檔/存檔/undo/redo | 基本圖檔操作 |
| `entity` | CRUD 幾何圖形 | line/circle/polyline/arc 等 |
| `layer` | 圖層管理 | 建立/凍結/鎖定/設定顏色 |
| `block` | Block 操作 | 插入 Block、讀寫屬性 |
| `annotation` | 標註操作 | 文字/標註/引線 |
| `pid` | P&ID 符號 | CTO library 符號庫 |
| `view` | 視圖操作 | 縮放/截圖 |
| `system` | 系統查詢 | 狀態查詢 + `execute_lisp` 擴展 |

### 6.3 mcp_dispatch.lsp — 通用 Dispatcher

autocad-mcp 提供的 `lisp-code/mcp_dispatch.lsp` 是 File IPC 的 AutoCAD 端核心：
- 讀取 `mcp_command_*.json` → dispatch 到對應 handler → 寫回 `mcp_result_*.json`
- 透過 `(c:mcp-dispatch)` 指令觸發（可由 PostMessageW 注入）
- 支援透過 `mcp:register-action` 擴展自訂 actions

### 6.4 ezdxf Backend（無頭模式）

autocad-mcp 支援三種 backend：
- **auto** — 自動偵測（優先 File IPC，fallback ezdxf）
- **file_ipc** — 透過 PostMessageW + JSON 與執行中的 AutoCAD 通訊
- **ezdxf** — 無需 AutoCAD 實例，直接讀寫 DXF/DWG 檔案

ezdxf backend 適用於 CI/CD 測試和批次處理場景。

---

## 7. File IPC 通訊協定

### 7.1 IPC 目錄

autocad-mcp 使用可配置的 IPC 目錄：
- 預設: `%TEMP%/autocad_mcp/`
- 環境變數: `AUTOCAD_MCP_IPC_DIR`

### 7.2 JSON Command/Result 格式

**Command（Python → AutoCAD）:**
```json
{
  "action": "odoo_extract_tables",
  "params": {
    "layout_name": null
  },
  "id": "cmd_1710000000_001"
}
```

**Result（AutoCAD → Python）:**
```json
{
  "success": true,
  "data": { ... },
  "id": "cmd_1710000000_001"
}
```

### 7.3 IPC 流程

```
呼叫端 (Python GUI 或 MCP Server)        AutoCAD (mcp_dispatch.lsp)
         │                                          │
    1. 寫 mcp_command_*.json 到 IPC 目錄             │
    2. 送 2×ESC (取消殘留命令)                        │
    3. PostMessageW(WM_CHAR) ──────────────────────→ │
       注入 "(c:mcp-dispatch)\n"                     │
       (Focus-Free，不搶焦點)                         │
         │                                    4. 讀 mcp_command JSON
         │                                    5. dispatch → 執行操作
         │                                    6. 寫 mcp_result_*.json
    7. 輪詢 mcp_result JSON ←────────────────────────│
    8. 讀取結果，刪除暫存檔                            │
```

### 7.4 通用 Actions（autocad-mcp 內建）

由 `mcp_dispatch.lsp` 直接處理，包括：`get_status`, `open_drawing`, `save_drawing`, `execute_lisp` 等。

### 7.5 Odoo 擴展 Actions（070_ob_mcp_dispatch.lsp 新增）

| Action | 說明 |
|--------|------|
| `odoo_extract_tables` | 收集所有 Layout TABLE + Block 資料 |
| `odoo_get_header_ids` | 收集所有 Layout 的 header_id |
| `odoo_write_ids` | 回寫 header_id + detail_id 到 TABLE |
| `odoo_get_block_attrs` | 讀取屬性 Block 的所有屬性 |
| `odoo_set_block_attrs` | 寫入屬性到所有 Layout 的 Block |
| `odoo_clear_ids` | 清除 TABLE ID |

### 7.6 超時與並發控制

| 機制 | 說明 |
|------|------|
| 可配置 timeout | `AUTOCAD_MCP_IPC_TIMEOUT`（1-300 秒，預設 10） |
| asyncio.Lock | Python 端防止並行 dispatch 競態 |
| PostMessageW(WM_CHAR) | Win32 API 送字元到 MDIClient 窗口，不搶焦點 |
| ESC 前置 | 2×ESC 取消殘留命令 |
| UTF-8 / cp1252 fallback | 自動處理編碼差異 |

---

## 8. MCP Server — AI 助手整合

### 8.1 概述

`mcp_server_autocad.py` 包裝 autocad-mcp 的 MCP server 並新增 Odoo 專用 tools，
提供 AI 助手（Claude / Gemini CLI）統一的自然語言操作介面。

### 8.2 Tool 分類

**8 個通用 AutoCAD Tools（來自 autocad-mcp）:**

| Tool | 功能 |
|------|------|
| `drawing` | 開檔/存檔/undo/redo |
| `entity` | CRUD 幾何圖形 |
| `layer` | 圖層管理 |
| `block` | Block 插入、屬性操作 |
| `annotation` | 文字/標註/引線 |
| `pid` | P&ID 符號 |
| `view` | 縮放/截圖 |
| `system` | 狀態查詢 + `execute_lisp` |

**5 個 Odoo 專用 Tools（我們擴展）:**

| Tool | 功能 |
|------|------|
| `odoo_extract` | 讀取所有 Layout TABLE + Block 資料 |
| `odoo_push_boq` | 收集 TABLE → 推送 Odoo → 回寫 ID |
| `odoo_create_pr` | 收集 header_ids → BOQ 轉 PR |
| `odoo_set_params` | 讀取 Odoo 選項 → 寫入 Block 屬性 |
| `odoo_status` | 檢查 Odoo 連線狀態 |

### 8.3 AI 使用範例

```
User: "幫我把目前圖檔的 BOQ 推送到 Odoo"

AI Assistant:
  1. 呼叫 system tool → get_status → 確認 AutoCAD 已連線
  2. 呼叫 odoo_status tool → 確認 Odoo 已連線
  3. 呼叫 odoo_extract tool → 收集所有 Layout TABLE 資料
  4. 呼叫 odoo_push_boq tool → 推送到 Odoo + 回寫 ID
  5. 回報結果: "已成功推送 3 個 Layout、15 筆明細到 Odoo BOQ"
```

### 8.4 execute_lisp 擴展能力

透過 autocad-mcp 的 `system` tool 的 `execute_lisp` action，AI 助手可以執行任意 AutoLISP 程式碼，實現無限擴展：

```
AI: system.execute_lisp("(table:get-all-layouts-data)")
→ 直接呼叫我們的 AutoLISP 函數，無需新增 MCP tool
```

---

## 9. Python GUI 層 — COM/IPC 模式切換

### 9.1 進入點 `odoo.py`

Phase 6 新增兩個命令列參數：

```bash
# 預設 COM 模式（現有行為不變）
python odoo.py

# 明確指定 COM 模式
python odoo.py --autocad-mode com

# IPC 模式（AutoCAD LT 2024+ 支援）
python odoo.py --autocad-mode ipc

# 啟動 MCP Server（搭配任一 AutoCAD 模式）
python odoo.py --autocad-mode ipc --mcp-autocad
```

`odoo.py` 解析參數後傳給 `FormMain`：

```python
# odoo.py (Phase 6 修改)
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--autocad-mode', choices=['com', 'ipc'], default='com')
parser.add_argument('--mcp-autocad', action='store_true')
# 保留舊參數相容性
parser.add_argument('--enable-mcp', action='store_true', help='(deprecated, use --mcp-autocad)')
args = parser.parse_args()

app = FormMain(odoo_connection, autocad_mode=args.autocad_mode,
               enable_mcp=args.mcp_autocad or args.enable_mcp)
```

### 9.2 FormMain 初始化流程

`forms/form_main_modern.py` 根據 `autocad_mode` 參數決定初始化路徑：

```python
class FormMain:
    def __init__(self, odoo_connection, autocad_mode='com', enable_mcp=False):
        self.autocad_mode = autocad_mode

    def init_utilities(self):
        # 統一介面：UtilAutoCADDispatcher
        self.dispatcher = UtilAutoCADDispatcher(mode=self.autocad_mode)

        # COM 模式需要 GUI Proxy（STA 線程安全）
        if self.autocad_mode == 'com':
            setup_gui_proxy_handlers(self.dispatcher)
            self.start_gui_proxy_timer()  # 100ms 輪詢

        # IPC 模式不需 GUI Proxy（File IPC 無 COM 線程問題）

        # 業務邏輯透過 dispatcher 透明切換
        self.push_boq = UtilPushToBoq(self.dispatcher, self.odoo_util)
        self.transfer_pr = UtilTransferBoqToPr(self.dispatcher, self.odoo_util)

        # MCP 管理器（全新，取代舊 MCPSSEManager）
        if self.enable_mcp:
            self.mcp_manager = MCPManager(self.dispatcher, self.odoo_util)
            self.mcp_manager.start()
```

### 9.3 UtilAutoCADDispatcher 設計

統一介面，持有兩個後端，依 mode 轉發呼叫：

```python
class UtilAutoCADDispatcher:
    def __init__(self, mode='com'):
        self.mode = mode
        self._com_backend = None   # UtilAutoCAD (lazy init)
        self._ipc_backend = None   # UtilAutoCADIPC (lazy init)

    @property
    def active_backend(self):
        if self.mode == 'com':
            if not self._com_backend:
                self._com_backend = UtilAutoCAD(...)
            return self._com_backend
        else:
            if not self._ipc_backend:
                self._ipc_backend = UtilAutoCADIPC(...)
            return self._ipc_backend

    def switch_mode(self, new_mode):
        """切換模式（可由 UI 設定區觸發）"""
        self.mode = new_mode

    # 統一介面方法 — 轉發到 active_backend
    def get_all_layouts_data(self): ...
    def get_all_header_ids(self): ...
    def write_ids(self, response): ...
    def get_block_attributes(self): ...
    def set_block_attributes(self, attrs, layout=None): ...
```

### 9.4 GUI Proxy 條件化

`utility/util_gui_proxy.py` 的 `setup_gui_proxy_handlers()` 加入模式判斷：

- **COM 模式**: 啟用 GUI Proxy（所有 COM 操作必須在 GUI 主線程 STA 執行）
- **IPC 模式**: bypass GUI Proxy（File IPC 是 process 間通訊，無 COM 線程問題）

### 9.5 MCP 層替換

**舊架構（刪除）:**
```
form_main_modern.py
  └→ MCPSSEManager(autocad_util, odoo_util)     [utility/util_mcp_sse_manager.py]  ← 刪除
       └→ import mcp_server_fastmcp              [mcp_server_fastmcp.py]            ← 刪除
       └→ StandardMCPSSEServer                    [utility/util_mcp_sse_server.py]   ← 刪除
```

**新架構（取代）:**
```
form_main_modern.py
  └→ MCPManager(dispatcher, odoo_util)            [utility/util_mcp_manager.py]     ← 全新
       └→ import mcp_server_autocad               [mcp_server_autocad.py]            ← 全新
            └→ 8 通用 tools (from autocad-mcp submodule)
            └→ 5 Odoo tools (我們擴展)
       └→ stdio / streamable-http transport        [取代舊 SSE 模式]
```

**刪除清單:**
- `utility/util_mcp_sse_manager.py` — 舊 SSE 管理器
- `utility/util_mcp_sse_server.py` — 舊 SSE transport
- `mcp_server_fastmcp.py` — 舊 7 tools MCP server

### 9.6 業務邏輯不需修改

`UtilPushToBoq` / `UtilTransferBoqToPr` 透過 dispatcher 統一介面操作 AutoCAD，
不需要知道底層是 COM 還是 IPC，模式切換對它們完全透明。

### 9.7 模式切換 UI

GUI 設定區域提供 RadioButton / OptionMenu，讓使用者在運行時切換 COM ↔ IPC：

```
┌─ AutoCAD Mode ────────────────────────┐
│  ● COM (Full AutoCAD, pywin32)        │
│  ○ IPC (AutoCAD LT 2024+, File IPC)  │
└───────────────────────────────────────┘
```

切換時呼叫 `dispatcher.switch_mode(new_mode)`，並更新狀態列顯示。

### 9.8 衝突避免

COM 模式和 IPC 模式不應同時操作 AutoCAD：
- 同一時間應只有一個 caller（Python COM 端 或 Python IPC 端）
- autocad-mcp 使用 `asyncio.Lock` 防止 Python 端並行 dispatch
- GUI 的模式切換是排他的（COM 或 IPC，不同時啟用）
