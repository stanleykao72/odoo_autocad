# CLAUDE.md

> **版本**: 6.0（雙模式 AutoCAD + 單一 MCP 伺服器）
> **最後更新**: 2026年7月31日

本檔提供 Claude Code 在此 repo 工作時的指引。**詳細架構請看 `docs/`，本檔只保留不易從程式碼推得的規則與常用指令。**

## Project Overview

Windows 桌面應用，橋接 **Odoo ERP** 與 **AutoCAD**，用於工程／營建流程：
從圖面擷取參數 → 產生 BOQ（工程量清單）→ 轉為 PR（請購單）並送入 Odoo。

## Architecture（重點）

雙後端透過共同的 ABC 抽象，業務邏輯不需判斷模式：

| 模式 | 後端 | 適用 AutoCAD | 通訊 | 執行緒策略 |
|------|------|--------------|------|-----------|
| **COM** | `utility/util_autocad.py` | Full AutoCAD 2019+ | pywin32 COM | **主執行緒**（STA 安全） |
| **IPC** | `utility/util_autocad_ipc.py` | LT 2024+ / Full | autocad-mcp File IPC | 背景執行緒 + ProgressDialog |

- `utility/autocad_backend_interface.py` — ABC，17 個抽象方法，兩個後端都必須實作
- `utility/util_autocad_dispatcher.py` — 統一入口，依模式路由；backend 延遲初始化
- **兩端 `get_layouts_values()` / `get_layouts_header_id_to_pr()` 都必須回傳 `{"all": [...]}`**，這是讓 `util_push_to_boq.py` 不需要模式判斷的關鍵契約
- `utility/util_gui_proxy.py` — queue 式代理，讓背景執行緒的請求回到 GUI 主執行緒執行 COM 操作。**僅 COM 模式使用**
- `utility/util_mcp_manager.py` — 在 GUI 行程內以背景 thread 跑 uvicorn，好讓 MCP 工具共用同一個 dispatcher／odoo_util 實例
- `mcp_server_autocad.py` — 單一 FastMCP 伺服器，共 13 個工具（上游 autocad-mcp 8 個 + 本專案 5 個 Odoo 工具）

完整說明見 `docs/ARCHITECTURE.md`、`docs/COM_IPC_ISOLATION.md`、`docs/WORKFLOWS.md`。

## Entry Point

```bash
python odoo.py                      # GUI（預設 COM 模式）
python odoo.py --autocad-mode ipc   # GUI（IPC 模式，AutoCAD LT）
python odoo.py --enable-mcp         # GUI + 自動啟動 MCP 伺服器
python odoo.py --mcp-autocad        # 純 MCP 伺服器（stdio，無 GUI）
```

## Development Commands

```bash
conda activate odoo_autocad
pip install -r requirements.txt
pip install -r requirements-windows.txt   # Windows 專用相依

# 測試（pytest.ini 已設定預設排除 live 測試）
pytest                                     # 全部單元／整合測試
pytest tests/integration -m autocad        # 需要 Full AutoCAD 正在執行
pytest tests/integration -m ipc            # 需要 AutoCAD + McpDispatch.vlx 已載入
pytest --cov=. --cov-report=term-missing

# 建置（一鍵：EXE + 安裝包，版本號自動取自 version.py）
build_and_package.bat      # 或 build_and_package.ps1
```

## 專案慣例（Claude 必須遵守）

### 版本號
`version.py` 的 `APP_VERSION` 是**唯一來源**。GUI 標題、log、建置腳本、Inno Setup 都由此取得。
每次交付都要 bump（格式 `MAJOR.MINOR.PATCH.BUILD`）。

### 機密處理
- `config/*.yaml` **一律不進 git**，只提交 `*.yaml.example` 範本
- `server*.yaml` 的 `url` query string 內含 API token，與 `token.yaml` 同等敏感
- **絕不把 token 寫進 log**：`UtilLog.safe_log_insert()` 會落地到 `logs/*.log`。需要顯示時用 `utility/util_secrets.py` 的 `mask_secret()`
- MCP 伺服器只綁 `127.0.0.1` — 13 個工具沒有任何認證機制

### 測試
- 遵循 TDD：先寫失敗測試 → 最小實作 → 重構
- **不要用 `sys.modules` 注入 mock 來模擬 COM**。`util_autocad.py` 在 import 時就綁定了 `pythoncom` / `client`，
  請改用 `patch.object(utility.util_autocad, 'pythoncom', ...)`，才不會受 import 順序影響
- **絕不直接改真實模組的屬性**（例如 `pythoncom.CoInitialize = MagicMock()`）— 會污染整個 session
- 需要真實 AutoCAD 的測試必須掛 `@pytest.mark.autocad` 或 `.ipc`

### 執行緒
- COM 操作**只能**在 GUI 主執行緒執行；長時間工作用 `_run_with_progress_main_thread()`
- IPC 操作可以在背景執行緒；用 `_run_with_progress_background()`
- 兩者由 `ModernFormMain._run_with_progress()` 依模式自動分流

## Key Files

```
odoo.py                              進入點（CLI + GUI launcher）
version.py                           APP_VERSION 唯一來源
mcp_server_autocad.py                MCP 伺服器（13 工具）

forms/form_main_modern.py            主視窗（CustomTkinter）
forms/form_autocad_param_enhanced.py 參數選擇表單（含 Odoo 產品查詢）

utility/autocad_backend_interface.py ABC 介面契約
utility/util_autocad.py              COM 後端
utility/util_autocad_ipc.py          IPC 後端
utility/util_autocad_dispatcher.py   COM/IPC 路由
utility/util_odoo.py                 Odoo REST API（Bravado/Swagger）
utility/util_push_to_boq.py          BOQ 推送流程
utility/util_transfer_boq_to_pr.py   PR 產生流程
utility/util_gui_proxy.py            GUI 主執行緒代理（COM only）
utility/util_mcp_manager.py          MCP 伺服器生命週期
utility/util_secrets.py              機密遮罩

autolisp/lisp/080_main.lsp           AutoLISP 進入點
autolisp/lisp/070_ob_mcp_dispatch.lsp 6 個 Odoo IPC actions
autolisp/build_vlx.lsp               VLX 建置（需 Full AutoCAD 的 VLIDE）
libs/autocad-mcp/                    git submodule (puran-water/autocad-mcp)
```

## 已知限制

- `utility/util_odoo.py` 的 `search_products()` / `push_boq_data()` 目前是回傳假資料的 stub
- `utility/util_autocad.py` 的 `scan_elements()` 是空實作（回傳空結果）
- autocad-mcp 需要 `pip install structlog`
- 不可 import 上游 autocad-mcp 的 tools 模組：其 `ToolResult = str | list` 型別別名會讓 FastMCP 產生 ForwardRef 錯誤
  （`mcp_server_autocad.py` 以注入 `builtins.ToolResult` 迴避）

## 文件

| 檔案 | 內容 |
|------|------|
| `docs/ARCHITECTURE.md` | 系統架構（Python 側） |
| `docs/API_REFERENCE.md` | Python / Odoo / MCP / AutoLISP API |
| `docs/FILE_MAP.md` | 完整檔案對照表 |
| `docs/WORKFLOWS.md` | 工作流程圖 |
| `docs/COM_IPC_ISOLATION.md` | COM/IPC 模式隔離架構 |
| `docs/COM_CONNECTION_STRATEGY.md` | COM 連線重試策略 |
| `docs/DEPLOYMENT_GUIDE.md` | 建置與部署 |
| `docs/INSTALL_GUIDE_USER.md` | 使用者安裝手冊 |
| `docs/CODE_SIGNING_GUIDE.md` / `ANTIVIRUS_SOLUTION.md` / `INTERNAL_CA_GUIDE.md` | 簽章與防毒 |
| `autolisp/docs/` | AutoLISP 架構、VLX 建置、LT 相容性研究 |
