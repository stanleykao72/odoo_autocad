# AutoCAD LT AutoLISP 支援研究報告

> 調查日期: 2026-03-06
> 目的: 評估將現有 AutoLISP + Python COM 架構改為純 AutoLISP 方案，以相容 AutoCAD LT

---

## 1. AutoCAD LT AutoLISP 支援現況

### 1.1 支援時間線

| 版本 | AutoLISP | DCL | VBA | COM API | ObjectARX |
|------|----------|-----|-----|---------|-----------|
| LT 2023 及以前 | 不支援 | 不支援 | 不支援 | 不支援 | 不支援 |
| **LT 2024+** | **支援** | **支援** | 不支援 | 不支援 | 不支援 |
| 完整版 AutoCAD | 支援 | 支援 | 支援 | 支援 | 支援 |

### 1.2 LT 2024+ 支援的功能

- AutoLISP 基本函數（.lsp, .fas, .vlx）
- DCL 對話框（.dcl）
- 大部分 `VL*` 函數（字串、清單操作）
- 大部分 `VLA*` / `VLAX*` 函數（操作**內建** AutoCAD 物件）
- `vlax-for` 遍歷集合
- TABLE 實體讀寫（`vla-GetText`, `vla-SetText`）
- Block 屬性讀寫（`vlax-invoke 'GetAttributes`）

### 1.3 LT 不支援的功能（關鍵限制）

| 函數 / 功能 | 影響 |
|-------------|------|
| `vlax-create-object` | 無法建立外部 COM 物件（XMLHTTP, ADODB, Python.ComServer） |
| `vlax-get-object` | 無法取得外部執行中的 COM 物件 |
| `vlax-get-or-create-object` | 同上 |
| `vlax-import-type-library` | 無法匯入型別庫 |
| `vla-GetInterfaceObject` | 無法取得 COM 介面 |
| 外部 COM 自動化 | pywin32 `GetActiveObject` / `CreateObject` 無法連接 LT |
| VBA / ObjectARX / .NET | 完全不支援 |
| VLIDE (IDE) | 無法在 LT 內除錯 |
| `vlisp-compile` | 無法在 LT 內編譯 |
| 3D 實體 / 表面 | 無法建立 3D 物件 |

---

## 2. 對現有 AutoLISP 檔案的影響

| 檔案 | LT 可用 | 原因 |
|------|---------|------|
| `json_util.lsp` | 可用 | 純字串操作 |
| `StripMtext v5-0b.lsp` | 可用 | 純內建物件操作 |
| `contract_product.lsp` | 可用 | DCL + 內建 VLA 操作 |
| `param_form.dcl` | 可用 | 純 DCL |
| `contract_product.dcl` | 可用 | 純 DCL |
| `read_csv.lsp` | 部分 | `LM:readcsv` 可用，ADODB 部分不可用 |
| **`call_python.lsp`** | **不可用** | 依賴 `vlax-get-or-create-object` |
| **`transfer_to_odoo.lsp`** | **不可用** | 所有 Odoo 整合透過 Python COM |

---

## 3. AutoLISP 直接 HTTP 通訊可行性

### 3.1 完整版 AutoCAD — 可行

```lisp
;; 使用 MSXML2.XMLHTTP（需要 vlax-create-object）
(setq http (vlax-create-object "MSXML2.XMLHTTP"))
(vlax-invoke-method http 'open "POST" url :vlax-false)
(vlax-invoke-method http 'setRequestHeader "Content-Type" "application/json")
(vlax-invoke-method http 'send json_body)
(setq response (vlax-get http 'responseText))
(vlax-release-object http)
```

### 3.2 AutoCAD LT — 不可行

`vlax-create-object` 在 LT 中永遠回傳 `nil`，無法建立 XMLHTTP 物件。
**結論：AutoLISP 在 LT 中無法直接發送 HTTP 請求。**

---

## 4. 替代方案評估

### 方案 A：startapp + curl + 檔案交換

```
AutoLISP → 寫 request.json
         → startapp curl (非同步)
         → 輪詢等待 response.json
         → 讀取結果
```

- 優點：不需額外安裝（Windows 10+ 內建 curl）
- 缺點：`startapp` 非同步、需輪詢、時序不精確、錯誤處理差

### 方案 B：startapp + MSHTA VBScript 橋接

```
AutoLISP → startapp mshta.exe vbscript:CreateObject("MSXML2.XMLHTTP")...
```

- 優點：可繞過 `vlax-create-object` 限制
- 缺點：字串轉義極複雜、維護困難、Windows Defender 可能攔截

### ~~方案 C：獨立 Python Bridge .exe（已棄用）~~

> **已棄用**: DCL 獨立模式已移除，由方案 D (IPC/MCP) 取代。

```
AutoLISP → 寫 request.json
         → startapp odoo_bridge.exe action request.json response.json
         → 輪詢等待 response.json
         → 讀取並解析結果
```

- 優點：邏輯清晰、好維護、可處理認證/重試/錯誤
- 缺點：需部署一個額外 .exe（可用 PyInstaller 打包）

### 方案 D：autocad-mcp submodule + Python 直連 + MCP（最終採用）

```
Python (GUI 或 MCP Server)
  → autocad-mcp File IPC → PostMessageW → AutoCAD 執行
  → Python 直接呼叫 Odoo API（不需 Bridge.exe）
  → AI 助手可透過 MCP 自然語言操作
```

- 優點：重用成熟開源方案、免費獲得 8 個通用 tools、三模式共存、AI 助手原生支援
- 缺點：依賴外部 submodule、需 AutoCAD LT 2024+

### 方案比較

| 項目 | A (curl) | B (MSHTA) | ~~C (Python Bridge)~~ | **D (autocad-mcp + MCP)** |
|------|----------|-----------|-------------------|--------------------------|
| LT 支援 | 可 | 可能被攔截 | ~~可~~ | **可 (LT 2024+)** |
| 維護性 | 中 | 低 | ~~高~~ | **高** |
| 錯誤處理 | 差 | 差 | ~~好~~ | **好** |
| 安全性 | 中 | 低 | ~~高~~ | **高** |
| 部署複雜度 | 低 | 低 | ~~中~~ | 中 |
| AI 助手支援 | 無 | 無 | ~~無~~ | **原生 MCP** |
| 通用 AutoCAD 工具 | 無 | 無 | ~~無~~ | **8 個（免費）** |
| DCL 獨立運作 | 需自建 | 需自建 | ~~可~~ | 不使用 DCL |
| 需 Bridge.exe | 否 | 否 | ~~是~~ | **否** |

---

## 5. 採用方案 D 的理由

> 方案 D 為最終採用的架構。方案 C（DCL 獨立模式）已棄用並移除。

1. **重用成熟開源方案** — autocad-mcp（puran-water/autocad-mcp）已驗證 File IPC 在 LT 2024+ 上可用，PostMessageW + JSON 機制穩定
2. **免費獲得 8 個通用 AutoCAD tools** — drawing/entity/layer/block/annotation/pid/view/system，加上 ezdxf 無頭後端
3. **完全重用現有 Python UI + Odoo 代碼** — `util_autocad_ipc.py` 包裝 autocad-mcp lib，`util_autocad_dispatcher.py` 統一 COM/IPC 介面
4. **雙模式共存** — COM（Full AutoCAD 使用者）、IPC（LT 使用者），Settings 頁面一鍵切換；MCP 模式供 AI 助手使用
5. **AI 助手可直接操作 AutoCAD** — puran-water/autocad-mcp 的核心功能，透過 MCP 協定自然語言驅動
6. **Dispatcher 模式** — `UtilAutoCADDispatcher` 統一 COM/IPC 介面，一鍵切換

---

## 6. 社群 IPC 方案調研 — puran-water/autocad-mcp

### 6.1 專案概述

| 項目 | 說明 |
|------|------|
| 名稱 | autocad-mcp |
| GitHub | https://github.com/puran-water/autocad-mcp |
| 版本 | v3.1+ |
| 語言 | Python + AutoLISP |
| 授權 | MIT |
| 功能 | 8 個 MCP tools + File IPC + ezdxf backend |
| LT 支援 | LT 2024+（需 AutoLISP 支援） |

### 6.2 File IPC 架構

autocad-mcp 使用 **PostMessageW + JSON 檔案交換** 實現 Focus-Free 的 IPC：

```
Python (MCP Server)                      AutoCAD (mcp_dispatch.lsp)
     │                                          │
1. 寫 mcp_command_*.json                        │
2. 送 2×ESC (取消殘留命令)                        │
3. PostMessageW(WM_CHAR) ─────────────────────→ │
   注入 "(c:mcp-dispatch)\n"                     │
   (不搶焦點)                                     │
     │                                    4. 讀 command JSON
     │                                    5. dispatch + 執行
     │                                    6. 寫 mcp_result_*.json
7. 輪詢 result JSON ←────────────────────────────│
8. 讀取結果，刪除暫存檔                            │
```

### 6.3 關鍵機制

| 機制 | 說明 |
|------|------|
| **PostMessageW(WM_CHAR)** | Win32 API 送字元到 MDIClient 窗口，不搶焦點 |
| **ESC 前置** | 2×ESC 取消殘留命令，確保命令列乾淨 |
| **UTF-8 / cp1252 fallback** | 自動處理 AutoCAD 的編碼差異 |
| **可配置 timeout** | `AUTOCAD_MCP_IPC_TIMEOUT`（1-300 秒，預設 10） |
| **asyncio.Lock** | 防止並行 dispatch 競態 |
| **auto/file_ipc/ezdxf backend** | 自動選擇後端（有 AutoCAD 用 IPC，無則用 ezdxf） |

### 6.4 execute_lisp 擴展能力

autocad-mcp 的 `system` tool 支援 `execute_lisp` action，可執行任意 AutoLISP 程式碼。這意味著：
- 我們的 `table_util.lsp` / `block_util.lsp` 函數可直接被呼叫
- 不需為每個操作都建立專用 MCP tool
- AI 助手具備無限擴展能力

### 6.5 LT 2024+ 相容性

autocad-mcp 已驗證在 AutoCAD LT 2024+ 上的可用性：
- `mcp_dispatch.lsp` 使用標準 AutoLISP 函數，不依賴 `vlax-create-object`
- PostMessageW 透過 Python ctypes 呼叫，不需 AutoCAD COM API
- 與我們的 LT 相容策略完全一致

---

## 參考來源

- [CAD Forum - LISP limitations in AutoCAD LT](https://www.cadforum.cz/en/limitations-of-the-lisp-language-autolisp-visuallisp-autocad-lt-tip13683)
- [CAD Forum - vlax-create-object workaround (Tip 14007)](https://www.cadforum.cz/en/workaround-for-the-missing-vlax-create-object-in-autocad-lt-tip14007)
- [Autodesk - AutoLISP support for AutoCAD LT](https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/Are-VBA-SCR-and-LISP-working-in-AutoCAD-LT.html)
- [Autodesk Blog - AutoCAD LT and AutoLISP](https://www.autodesk.com/blogs/autocad/autocad-lt-2024-autolisp/)
- [AutoCAD DevBlog - Entitlement API with LISP (XMLHTTP example)](https://adndevblog.typepad.com/autocad/2022/05/using-entitlement-api-with-lisp.html)
- [Autodesk Forum - REST POST API with LISP](https://forums.autodesk.com/t5/visual-lisp-autolisp-and-general/lisp-function-with-rest-post-api-is-it-doable/td-p/13282194)
