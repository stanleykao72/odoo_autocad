# AutoLISP IPC/MCP 模式 — 實作計畫

> 版本: 3.0 (DCL 獨立模式已移除)
> 日期: 2026-03-10
> 狀態: Phase 1-5 已移除（DCL 獨立模式）；Phase 6 計畫中（13 項目: 6 新增 + 4 修改 + 3 刪除）

---

## 1. 背景與目標

基於 `puran-water/autocad-mcp` 開源專案（MIT 授權），實作 IPC 和 MCP 兩種操作模式，
與現有 COM 模式共存，形成雙模式架構（COM / IPC+MCP）。

### 核心需求

- **LT 相容**: 所有 `.lsp` 檔不含 `vlax-create-object` / `vlax-get-or-create-object`（例外：`040_strip_mtext.lsp` 使用 `vl-catch-all-apply` 包裹 + LT fallback）
- **IPC Mode**: Python GUI 透過 File IPC（autocad-mcp）驅動 AutoCAD LT 2024+
- **MCP Mode**: AI 助手透過 MCP Server + File IPC 驅動 AutoCAD
- **COM Mode 不變**: 現有 Python GUI + COM 直連 Full AutoCAD 完全不受影響

### 歷史

Phase 1-5 曾實作 DCL 對話框 + Python Bridge.exe 的獨立模式架構（方案 C）。
該模式已被 Phase 6 的 IPC/MCP 模式（方案 D）取代，相關程式碼已刪除。
舊版檔案保留在 `legacy/` 目錄供歷史參考。

---

## 2. 保留的基礎模組（Phase 3 實作，IPC/MCP 共用）

以下模組在 Phase 3 實作完成，IPC/MCP 模式繼續使用：

| # | 檔案 | 說明 | 狀態 |
|---|------|------|------|
| 1 | `lisp/030_config.lsp` | 路徑常數、YAML 設定讀取 | DONE |
| 2 | `lisp/010_json_util.lsp` | JSON 解析/序列化 | DONE |
| 3 | `lisp/020_file_util.lsp` | 檔案讀寫 + 唯一檔名 | DONE |
| 4 | `lisp/060_table_util.lsp` | TABLE 遍歷 + 資料收集 + ID 回寫 + 清除 | DONE |
| 5 | `lisp/050_block_util.lsp` | Block 屬性讀寫 + 搜尋 | DONE |
| 6 | `lisp/040_strip_mtext.lsp` | MText 格式清除（RegExp + LT fallback） | DONE |

---

## 3. Phase 6: autocad-mcp Submodule + IPC/MCP Mode — IN PROGRESS

### 背景

基於 `puran-water/autocad-mcp` 開源專案（MIT 授權），新增 IPC 和 MCP 兩種操作模式，
與現有 COM 模式共存。詳見 [ARCHITECTURE.md](ARCHITECTURE.md) Section 1。

### 實作檔案 — 新增 (6 項)

| # | 檔案 | 動作 | 內容 | 狀態 |
|---|------|------|------|------|
| 1 | `libs/autocad-mcp/` | 新增 | git submodule (puran-water/autocad-mcp) | DONE |
| 2 | `lisp/070_ob_mcp_dispatch.lsp` | 新增 | Odoo 擴展 dispatcher（載入 autocad-mcp `mcp_dispatch.lsp` + 註冊 Odoo actions） | DONE |
| 3 | `utility/util_autocad_ipc.py` | 新增 | Python File IPC client（包裝 autocad-mcp lib，提供與 `util_autocad.py` 相似的介面） | DONE |
| 4 | `utility/util_autocad_dispatcher.py` | 新增 | COM/IPC 模式切換（統一介面） | DONE |
| 5 | `mcp_server_autocad.py` | 新增 | MCP Server（包裝 autocad-mcp 的 8 通用 tools + 新增 5 個 Odoo tools = 13 tools） | DONE |
| 6 | `lisp/080_main.lsp` | 修改 | 新增 `OB:MCP-DISPATCH` 指令，載入 `070_ob_mcp_dispatch.lsp`（DCL 相關已移除） | DONE |

### 實作檔案 — 修改現有 (4 項)

| # | 檔案 | 動作 | 內容 | 狀態 |
|---|------|------|------|------|
| 7 | `forms/form_main_modern.py` | 修改 | 改用 `UtilAutoCADDispatcher`；COM 模式才初始化 GUI proxy；改用 `MCPManager`（取代 `MCPSSEManager`）；新增模式切換 UI | DONE |
| 8 | `utility/util_gui_proxy.py` | 修改 | `setup_gui_proxy_handlers()` 加入模式判斷：COM 走 proxy，IPC bypass | DONE (在 form_main_modern.py 中判斷) |
| 9 | `odoo.py` | 修改 | 新增 `--autocad-mode com|ipc`、`--mcp-autocad` 參數；傳 mode 給 FormMain | DONE |
| 10 | `utility/util_mcp_manager.py` | 新增 | 全新 MCP 管理器，整合 `mcp_server_autocad.py`，取代舊 SSE 模式 | DONE |

### 實作檔案 — 刪除 (3 項，SSE 舊模式)

| # | 檔案 | 動作 | 說明 | 狀態 |
|---|------|------|------|------|
| 11 | `utility/util_mcp_sse_manager.py` | 刪除 | 由 `util_mcp_manager.py` (#10) 取代 | TODO |
| 12 | `utility/util_mcp_sse_server.py` | 刪除 | 由 `mcp_server_autocad.py` (#5) 內建 transport 取代 | TODO |
| 13 | `mcp_server_fastmcp.py` | 刪除 | 由 `mcp_server_autocad.py` (#5) 取代（7→13 tools） | TODO |

### Phase 6 驗證

| # | 驗證項目 | 說明 |
|---|----------|------|
| 1 | Submodule 驗證 | `git submodule update --init` → `libs/autocad-mcp/` 存在 |
| 2 | 通用 tool 測試 | MCP Server → `system` tool → `get_status` → 回傳 AutoCAD 資訊 |
| 3 | Odoo tool 測試 | MCP Server → `odoo_push_boq` → 收集 TABLE → 推送 Odoo → 回寫 ID |
| 4 | IPC 模式測試 | Python GUI Settings 選 IPC → Push BOQ → 結果與 COM 模式一致 |
| 5 | AI 測試 | Claude/Gemini CLI → 自然語言操作 AutoCAD → 驗證結果正確 |
| 6 | GUI COM 不變 | `python odoo.py` → 預設 COM 模式，現有功能正常 |
| 7 | GUI IPC 模式 | `python odoo.py --autocad-mode ipc` → 狀態顯示 "IPC Mode"，Push BOQ 結果一致 |
| 8 | MCP 替換驗證 | GUI 啟動後 MCP 可用（非 SSE），13 tools 可列出（取代舊 7 tools） |

---

## 4. 關鍵參考檔案

| 檔案 | 用途 |
|------|------|
| `autolisp/legacy/transfer_to_odoo.lsp` | TABLE 操作邏輯藍本（historical reference） |
| `autolisp/legacy/contract_product.lsp` | Block 屬性模式（historical reference） |
| `autolisp/legacy/json_util.lsp` | JSON 解析器（直接沿用） |
| `autolisp/docs/ARCHITECTURE.md` | 雙模式架構設計（v4.0） |
| `autolisp/docs/AUTOCAD_LT_RESEARCH.md` | LT 相容性研究 + 方案 D 評估 |
| `libs/autocad-mcp/` | puran-water/autocad-mcp submodule |
