# AutoLISP 安裝手冊

> **版本**: 1.1
> **適用**: AutoCAD LT 2024+ / Full AutoCAD 2024+
> **最後更新**: 2026-03-10

---

## 1. 系統需求

| 項目 | 最低要求 |
|------|----------|
| AutoCAD | AutoCAD LT 2024 或 Full AutoCAD 2024 以上 (Windows) |
| 作業系統 | Windows 10 / 11 |
| Python 端 | Python 3.10+（主機端已安裝 odoo-autocad-integration） |

> AutoCAD LT 2024 起才支援 AutoLISP。Mac 版 LT **不**支援 AutoLISP。

---

## 2. 檔案清單

### 2.1 Odoo 模組（8 個 .lsp）

位於 `autolisp/lisp/`，檔名前綴數字代表編譯/載入順序：

| 順序 | 檔案 | 說明 |
|------|------|------|
| 010 | `010_json_util.lsp` | JSON 解析器（純 AutoLISP） |
| 020 | `020_file_util.lsp` | 檔案 I/O 工具 |
| 030 | `030_config.lsp` | 路徑與設定管理 |
| 040 | `040_strip_mtext.lsp` | MText 格式清理 |
| 050 | `050_block_util.lsp` | Block / Attribute 操作 |
| 060 | `060_table_util.lsp` | TABLE 實體操作 |
| 070 | `070_ob_mcp_dispatch.lsp` | Odoo 擴展指令（6 個 action） |
| 080 | `080_main.lsp` | **入口點** — 載入上述模組，註冊 `OB:MCP-DISPATCH` |

### 2.2 autocad-mcp 基底 dispatcher（1 個 .lsp）

位於 `libs/autocad-mcp/lisp-code/`：

| 檔案 | 說明 |
|------|------|
| `mcp_dispatch.lsp` | 70+ 通用指令 dispatcher（IPC 通訊核心） |

---

## 3. 安裝方式

### 方式 A：散裝 .lsp（開發環境推薦）

最簡單的方式，適合開發測試。

#### 步驟 1：設定 AutoCAD 搜尋路徑

1. AutoCAD 命令列輸入 `OPTIONS`
2. 選擇 **Files** 分頁
3. 展開 **Support File Search Path**
4. 點 **Add**，加入以下兩個路徑：
   ```
   C:\odoo\autocad_source\autolisp\lisp
   C:\odoo\autocad_source\libs\autocad-mcp\lisp-code
   ```
5. 按 **OK** 確認

#### 步驟 2：載入 080_main.lsp

AutoCAD 命令列輸入：
```
APPLOAD
```
瀏覽到 `C:\odoo\autocad_source\autolisp\lisp\080_main.lsp`，點 **Load**。

成功載入會看到：
```
[OB] ========================================
[OB] Odoo-AutoCAD Integration (AutoLISP)
[OB] Loading modules...
[OB] ========================================
[OB] 030_config.lsp loaded
=== MCP Dispatch v3.1 loaded ===
[OB] mcp_dispatch.lsp loaded from file
[OB] ob_mcp_dispatch.lsp loaded
[OB] ========================================
[OB] All modules loaded successfully!
[OB] Commands:
[OB]   OB:MCP-DISPATCH - IPC/MCP dispatcher
[OB] ========================================
```

#### 步驟 3：設定自動載入（可選）

在 APPLOAD 對話框中，將 `080_main.lsp` 加入 **Startup Suite**。
之後每次開啟 AutoCAD 都會自動載入。

---

### 方式 B：打包為雙 VLX（部署推薦）

將所有 .lsp 編譯為 **兩個 VLX** 檔案，方便部署分發。

> **注意**: 建置 VLX 需要 **Full AutoCAD**（LT 沒有 VLIDE 編譯器），
> 但產出的 VLX 可在 LT 2024+ 上載入執行。

#### 產出物

| VLX | 包含的 .lsp | 用途 |
|-----|------------|------|
| `McpDispatch.vlx` | `mcp_dispatch.lsp` + `attribute_tools.lsp` | autocad-mcp IPC 通訊核心 |
| `OdooAutoCAD.vlx` | 8 個 Odoo .lsp（010~080，見 Section 2.1） | Odoo 整合模組 |

> `mcp_dispatch.lsp` 是完全自足的（JSON 解析、IPC 通訊全部內建），
> 可以獨立打包為 VLX。`pid_tools` 等可選功能找不到時會 graceful fallback。

#### 步驟 1：開啟 VLIDE

在 **Full AutoCAD** 命令列輸入：
```
VLIDE
```

#### 步驟 2：建置 McpDispatch.vlx

1. File → Project → New Project → 專案名稱：`McpDispatch`
2. 加入源碼：
   ```
   1. libs/autocad-mcp/lisp-code/attribute_tools.lsp
   2. libs/autocad-mcp/lisp-code/mcp_dispatch.lsp   ← Entry Point
   ```
3. Project → Build Project
4. 輸出檔：`autolisp/dist/McpDispatch.vlx`

#### 步驟 3：建置 OdooAutoCAD.vlx

1. File → Project → New Project → 專案名稱：`OdooAutoCAD`
2. 加入源碼（順序重要 — 依賴關係，按檔名前綴排序）：
   ```
   1. lisp/010_json_util.lsp
   2. lisp/020_file_util.lsp
   3. lisp/030_config.lsp
   4. lisp/040_strip_mtext.lsp
   5. lisp/050_block_util.lsp
   6. lisp/060_table_util.lsp
   7. lisp/070_ob_mcp_dispatch.lsp
   8. lisp/080_main.lsp              ← Entry Point
   ```
3. Project → Build Project
4. 輸出檔：`autolisp/dist/OdooAutoCAD.vlx`

#### 步驟 4：載入順序

**重要**：必須先載入 `McpDispatch.vlx` 再載入 `OdooAutoCAD.vlx`，
因為 `070_ob_mcp_dispatch.lsp` 會引用 `mcp-dispatch-command` 函數。

```
APPLOAD → McpDispatch.vlx   (先)
APPLOAD → OdooAutoCAD.vlx   (後)
```

#### 步驟 5：部署

部署只需提供 **2 個檔案**：

```
McpDispatch.vlx      ← IPC 核心（1378 行 → ~60KB VLX）
OdooAutoCAD.vlx      ← Odoo 模組（~50KB VLX）
```

#### 步驟 6：使用者安裝

1. 將兩個 `.vlx` 複製到目標機器同一目錄
2. 在 AutoCAD `OPTIONS` → Support File Search Path 加入該目錄
3. `APPLOAD` → 先載入 `McpDispatch.vlx`，再載入 `OdooAutoCAD.vlx`
4. 兩個都加入 Startup Suite（注意 McpDispatch 排在前面）

---

### 方式 C：自動化建置腳本

檔案位置：`autolisp/build_vlx.lsp`

在 Full AutoCAD 命令列執行：
```
(load "C:/odoo/autocad_source/autolisp/build_vlx.lsp")
BUILD-ALL
```

輸出：
```
autolisp/dist/McpDispatch.vlx   (IPC core)
autolisp/dist/OdooAutoCAD.vlx   (Odoo modules)
```

也可以單獨建置：
- `BUILD-MCP-DISPATCH` — 只建置 McpDispatch.vlx
- `BUILD-ODOO-VLX` — 只建置 OdooAutoCAD.vlx

---

## 4. 驗證安裝

載入完成後，在 AutoCAD 命令列測試：

```
OB:MCP-DISPATCH
```

若正常運作，Python 端（GUI 切換至 IPC 模式）發送指令後，
AutoCAD 會從 `%TEMP%` 讀取 JSON 指令、執行、回寫結果。

### 驗證通訊流程

```
Python (IPC模式)
  ↓ 寫入 %TEMP%\autocad_mcp_cmd_{id}.json
  ↓ PostMessageW 發送鍵盤訊號
AutoCAD LT
  ↓ mcp_dispatch.lsp 讀取 JSON
  ↓ 執行指令（繪圖/取值/寫入）
  ↓ 寫入 %TEMP%\autocad_mcp_result_{id}.json
Python
  ↑ 輪詢讀取結果 JSON
```

---

## 5. 疑難排解

### 載入失敗：找不到 mcp_dispatch.lsp

```
[OB] Warning: mcp_dispatch.lsp not found — IPC/MCP mode unavailable
```

**解法**：確認 `OPTIONS` → Support File Search Path 已包含
`libs/autocad-mcp/lisp-code/` 目錄。

### 載入失敗：找不到模組

```
[OB] Warning: Could not find 010_json_util.lsp
```

**解法**：確認 Support File Search Path 已包含 `autolisp/lisp/` 目錄。

### VLX 載入安全性警告

AutoCAD 可能顯示「此檔案未經簽署」警告。

**解法**：
1. 命令列輸入 `SECURELOAD`，設為 `0`（信任所有）或 `1`（提示）
2. 或對 VLX 進行數位簽章（參見 `doc/CODE_SIGNING_GUIDE.md`）

### IPC 無回應

**確認項目**：
1. AutoCAD 視窗未被最小化（PostMessageW 需要視窗可見）
2. `OB:MCP-DISPATCH` 指令已載入（命令列輸入測試）
3. `%TEMP%\odoo_bridge\` 目錄有寫入權限

---

## 6. 部署檢查清單

### 散裝 .lsp 部署

- [ ] 複製 `autolisp/lisp/` 下 8 個 .lsp（010~080）至目標機器
- [ ] 複製 `libs/autocad-mcp/lisp-code/mcp_dispatch.lsp` 至目標機器
- [ ] AutoCAD `OPTIONS` 設定 Support File Search Path
- [ ] `APPLOAD` 載入 `080_main.lsp` → 確認所有模組載入成功
- [ ] 加入 Startup Suite 實現自動載入
- [ ] 測試 `OB:MCP-DISPATCH` 指令

### 雙 VLX 部署

- [ ] 在 Full AutoCAD 執行 `BUILD-ALL` 建置兩個 VLX
- [ ] 複製 `McpDispatch.vlx` + `OdooAutoCAD.vlx` 至目標機器
- [ ] AutoCAD `OPTIONS` 設定 Support File Search Path
- [ ] `APPLOAD` 先載入 `McpDispatch.vlx` → 確認 `MCP Dispatch v3.1 loaded`
- [ ] `APPLOAD` 再載入 `OdooAutoCAD.vlx` → 確認 `All modules loaded`
- [ ] 兩個都加入 Startup Suite（McpDispatch 排前面）
- [ ] 測試 `OB:MCP-DISPATCH` 指令

---

## 7. 相關文件

| 文件 | 說明 |
|------|------|
| [ARCHITECTURE.md](ARCHITECTURE.md) | 系統架構設計 |
| [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md) | 實作計畫與進度 |
| [AUTOCAD_LT_RESEARCH.md](AUTOCAD_LT_RESEARCH.md) | AutoCAD LT 相容性研究 |
