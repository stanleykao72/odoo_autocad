# AutoLISP + Python Bridge 架構設計

> 版本: 2.0 (實作完成版)
> 日期: 2026-03-06
> 架構: 方案 C — Python Bridge .exe + 檔案交換
> 實作計畫: 見 [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md)

---

## 1. 架構概覽

```
+----------------------------------+      +------------------+      +-----------+
|  AutoCAD (LT or Full)            |      |  odoo_bridge.exe |      |  Odoo ERP |
|                                  |      |  (Python)        |      |           |
|  +----------------------------+  |      |                  |      |           |
|  | AutoLISP                   |  |      |  - HTTP Client   |      |           |
|  |                            |  |      |  - Auth Manager  |      |           |
|  |  1. 收集 TABLE/Block 資料  |  |      |  - JSON 解析     |      |           |
|  |  2. 寫 request.json        |--|----->|  - 錯誤處理      |----->|  Swagger  |
|  |  3. startapp bridge.exe    |  |      |  - 寫 response   |      |  API v2   |
|  |  4. 輪詢 response.json     |  |      |                  |<-----|           |
|  |  5. 讀取並處理結果         |<-|------|  - 寫 resp.json  |      |           |
|  |  6. 回寫 ID 到 TABLE       |  |      |                  |      |           |
|  +----------------------------+  |      +------------------+      +-----------+
|                                  |
|  +----------------------------+  |
|  | DCL 對話框 (While 迴圈)    |  |
|  |  - 主選單 (main_menu.dcl)  |  |
|  |  - 參數選擇 (param_form)   |  |
|  |  - 設定 (config.dcl)       |  |
|  |  - 結果顯示 (result.dcl)   |  |
|  +----------------------------+  |
+----------------------------------+

通訊介面: JSON 檔案 (temp 目錄, 唯一檔名, 用完即刪)
```

---

## 2. 目錄結構（已實作）

```
autolisp/
├── README.md                     # 安裝與使用說明
├── docs/                         # 文件
│   ├── ARCHITECTURE.md           # 本文件（架構 + 模組規格）
│   ├── IMPLEMENTATION_PLAN.md    # 實作計畫 + 進度追蹤
│   └── AUTOCAD_LT_RESEARCH.md   # LT 支援研究報告
│
├── lisp/                         # AutoLISP 原始碼（LT 相容）
│   ├── main.lsp                  # 進入點，載入所有模組，定義 8 個使用者指令
│   ├── main_menu.lsp             # 主選單 while 迴圈 + action dispatch
│   ├── config.lsp                # 路徑常數、INI 讀寫
│   ├── json_util.lsp             # JSON 解析/序列化（diegomcas 版本）
│   ├── file_util.lsp             # 檔案讀寫 + 唯一檔名 + 刪除
│   ├── odoo_bridge.lsp           # Bridge 通訊層（寫 req / startapp / 輪詢 / 讀 resp）
│   ├── table_util.lsp            # TABLE 實體讀寫（遍歷 Layout, 讀取/回寫 ID）
│   ├── block_util.lsp            # Block 屬性讀寫（get/set attribute）
│   ├── param_form.lsp            # 參數選擇表單邏輯（載入選項、使用者選擇、寫入屬性）
│   └── strip_mtext.lsp           # MText 格式清除（RegExp + LT 純字串 fallback）
│
├── dcl/                          # DCL 對話框定義
│   ├── main_menu.dcl             # 主選單（7 按鈕 + 狀態列）
│   ├── param_form.dcl            # 參數表單（Product 關鍵字搜尋 + 6 下拉選單）
│   ├── config.dcl                # 連線設定（URL, DB, username, password）
│   └── result.dcl                # 結果顯示（title + 5-line message + OK）
│
├── bridge/                       # Python Bridge 原始碼
│   ├── odoo_bridge.py            # 主程式（CLI 入口，action dispatch）
│   ├── odoo_client.py            # Odoo Swagger/REST API 客戶端
│   ├── auth.py                   # 認證管理（BasicAuth）
│   ├── config.py                 # INI 設定檔讀取（configparser）
│   ├── requirements.txt          # Python 依賴（requests, configparser）
│   └── build.bat                 # PyInstaller 打包腳本
│
├── config/                       # 設定檔
│   ├── bridge.ini                # Bridge 設定（.gitignore 中）
│   └── bridge.ini.example        # 設定檔範例
│
├── legacy/                       # 舊版檔案（參考用，不再直接使用）
│   ├── call_python.lsp           # 舊版 Python COM 呼叫
│   ├── transfer_to_odoo.lsp      # 舊版 Odoo 整合（COM 方式）
│   ├── contract_product.lsp      # 舊版參數選擇
│   ├── contract_product.dcl      # 舊版 DCL
│   ├── param_form.dcl            # 舊版 DCL
│   ├── json_util.lsp             # JSON 工具（已移入 lisp/）
│   ├── read_csv.lsp              # CSV 工具（已棄用，改由 Bridge API 取資料）
│   └── StripMtext v5-0b.lsp      # MText 工具（已移入 lisp/ 並加 LT fallback）
│
└── dist/                         # 建置輸出（.gitignore）
    └── odoo_bridge.exe           # 打包後的 Bridge 執行檔
```

---

## 3. 模組規格

### 3.1 lisp/config.lsp — 設定管理

**全域變數:**

| 變數 | 類型 | 說明 |
|------|------|------|
| `*ob:root*` | string | autolisp 根目錄路徑 |
| `*ob:lisp-dir*` | string | lisp/ 目錄路徑 |
| `*ob:dcl-dir*` | string | dcl/ 目錄路徑 |
| `*ob:bridge-dir*` | string | bridge/ 目錄路徑 |
| `*ob:config-dir*` | string | config/ 目錄路徑 |
| `*ob:dist-dir*` | string | dist/ 目錄路徑 |
| `*ob:bridge-exe*` | string/nil | Bridge 執行檔完整路徑（nil 表示 dev mode） |
| `*ob:ini-file*` | string | bridge.ini 路徑 |
| `*ob:temp-dir*` | string | JSON 暫存目錄（%TEMP%/odoo_bridge/） |
| `*ob:odoo-connected*` | T/nil | Odoo 連線狀態 |
| `*ob:bridge-timeout*` | int | Bridge 輪詢超時（ms，預設 30000） |
| `*ob:odoo-url*` | string | Odoo 伺服器 URL |
| `*ob:odoo-db*` | string | Odoo 資料庫名稱 |
| `*ob:odoo-user*` | string | 使用者名稱 |
| `*ob:odoo-pass*` | string | 密碼/API Key |

**函數:**

| 函數 | 說明 |
|------|------|
| `(config:init)` | 初始化所有路徑，偵測 bridge.exe 位置 |
| `(config:load-ini)` | 讀取 bridge.ini 到全域變數 |
| `(config:save-ini)` | 將目前設定寫回 bridge.ini |
| `(ini:read-file path)` | 解析 INI 檔，回傳 `(("section.key" . "value") ...)` |
| `(ini:get data section key)` | 從 INI 資料取值 |
| `(ini:write-file path data)` | 寫回 INI 檔 |

---

### 3.2 lisp/json_util.lsp — JSON 解析/序列化

基於 diegomcas 版本，修改 `value_to_string` 加入 `json:escape` 呼叫。

| 函數 | 說明 |
|------|------|
| `(dmc:json:json_to_list json quote_array)` | JSON 字串 → assoc list |
| `(dmc:json:list_to_json lst)` | assoc list → JSON 字串 |
| `(dmc:json:str_replace str patt repl_to)` | 字串全域替換 |

---

### 3.3 lisp/file_util.lsp — 檔案 I/O

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
| `(dcl:load filename)` | 載入 DCL（VLX 內嵌優先，fallback 外部檔案） |

---

### 3.4 lisp/odoo_bridge.lsp — Bridge 通訊層

**核心函數:**

| 函數 | 說明 |
|------|------|
| `(bridge:call action payload)` | 寫 req.json → startapp → 輪詢 → 讀 resp.json → 清理 |
| `(bridge:launch action req resp)` | 啟動 bridge exe 或 python（dev mode） |
| `(bridge:wait-for-response resp timeout)` | 每 500ms 輪詢 findfile，超時回傳錯誤 |

**回應解析:**

| 函數 | 說明 |
|------|------|
| `(bridge:success-p response)` | 檢查 success 欄位 |
| `(bridge:get-data response)` | 取得 data 欄位 |
| `(bridge:get-message response)` | 取得 message 欄位 |
| `(bridge:get-error response)` | 取得 error_code 欄位 |

**API Wrappers（7 個）:**

| 函數 | Bridge Action | 說明 |
|------|--------------|------|
| `(bridge:test-connection)` | `test_connection` | 測試連線 |
| `(bridge:get-project pr-no)` | `get_project` | 取得專案資訊 |
| `(bridge:get-products)` | `get_products` | 取得產品清單 |
| `(bridge:get-setup)` | `get_setup` | 取得設定值選項 |
| `(bridge:get-colors project-id)` | `get_colors` | 取得顏色清單 |
| `(bridge:import-to-boq data)` | `import_to_boq` | 推送 BOQ |
| `(bridge:boq-to-pr header-ids)` | `boq_to_pr` | BOQ 轉 PR |

---

### 3.5 lisp/table_util.lsp — TABLE 實體操作

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
| `(table:update-ids-from-response response)` | 從 bridge 回應批量回寫 ID |

**清除:**

| 函數 | 說明 |
|------|------|
| `(table:clear-ids ename)` | 清除單一 TABLE 的所有 ID |
| `(table:clear-current-layout-ids)` | 清除目前 Layout |
| `(table:clear-all-layouts-ids)` | 清除所有 Layout |

---

### 3.6 lisp/block_util.lsp — Block 屬性操作

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

### 3.7 lisp/strip_mtext.lsp — MText 格式清除

| 函數 | 說明 |
|------|------|
| `(strip:unformat str)` | 統一入口：自動選擇 RegExp 或 fallback |
| `(LM:UnFormat str mtx)` | Lee Mac 版（VBScript.RegExp，Full AutoCAD） |
| `(strip:unformat-simple str)` | LT fallback（純 `vl-string-subst` 鏈式替換） |

**LT fallback 處理的格式碼:**
`\P`(段落), `\f`/`\F`(字型), `\C`/`\c`(顏色), `\H`(高度), `\W`(寬度), `\T`(追蹤), `\Q`(傾斜), `\A`(對齊), `\L`/`\l`(底線), `\O`/`\o`(刪除線), `\~`(不分行空格), `{}`(群組)

---

### 3.8 lisp/param_form.lsp — 參數選擇表單

| 函數 | 說明 |
|------|------|
| `(param:parse-products data)` | 解析產品 → (ids names uoms) |
| `(param:parse-setup data)` | 解析設定 → (catalogs specs ops surfs) |
| `(param:parse-colors data)` | 解析顏色 → (names nos)，加 "No Color" |
| `(param:load-lov-data)` | 從 Odoo 載入所有 LOV 資料（3 次 bridge call） |
| `(param:filter-products keyword)` | 以關鍵字過濾產品清單並更新 list_box |
| `(param:get-selected-product-index str)` | 將 list_box 選項對應回原始產品索引 |
| `(param:show-form lov-list)` | 顯示 DCL 表單，回傳選擇結果 |
| `(param:apply-to-block block selections)` | 將選擇結果寫入 Block 屬性 |

---

### 3.9 lisp/main_menu.lsp — 主選單邏輯

| 函數 | 說明 |
|------|------|
| `(menu:show)` | 主選單 while 迴圈（7 個 action code） |
| `(menu:show-result title msg)` | 結果對話框（fallback 到 alert） |
| `(menu:do-connect)` | 測試 Odoo 連線 |
| `(menu:do-config)` | 設定對話框 → 存 bridge.ini |
| `(menu:do-set-params)` | 載入 LOV → 找 Block → 顯示表單 → 寫入 |
| `(menu:do-push-boq)` | 收集 TABLE → bridge → 回寫 ID |
| `(menu:do-create-pr)` | 收集 header_ids → bridge → PR |
| `(menu:do-clear-current)` | 清除目前 Layout TABLE ID |
| `(menu:do-clear-all)` | 清除所有 Layout TABLE ID |

---

### 3.10 lisp/main.lsp — 進入點

**載入順序:**
1. json_util.lsp
2. file_util.lsp
3. config.lsp
4. strip_mtext.lsp
5. block_util.lsp
6. table_util.lsp
7. odoo_bridge.lsp
8. param_form.lsp
9. main_menu.lsp

**使用者指令（8 個）:**

| 指令 | 函數 | 說明 |
|------|------|------|
| `OB:MENU` | `menu:show` | 開啟主選單 |
| `OB:CONNECT` | `menu:do-connect` | 測試 Odoo 連線 |
| `OB:CONFIG` | `menu:do-config` | 連線設定 |
| `OB:SET-PARAMS` | `menu:do-set-params` | 參數選擇表單 |
| `OB:PUSH-BOQ` | `menu:do-push-boq` | 推送 BOQ |
| `OB:CREATE-PR` | `menu:do-create-pr` | BOQ 轉 PR |
| `OB:CLEAR-IDS` | `menu:do-clear-current` | 清除目前 Layout ID |
| `OB:CLEAR-ALL-IDS` | `menu:do-clear-all` | 清除所有 Layout ID |

---

## 4. Python Bridge 規格

### 4.1 bridge/odoo_bridge.py — CLI 入口

```
用法: odoo_bridge.exe <action> <request.json> <response.json> [--config bridge.ini]

動作:
  test_connection     測試 Odoo 連線
  get_project         取得專案 (需 pr_no 參數)
  get_products        取得產品清單
  get_setup           取得設定值
  get_colors          取得顏色清單
  import_to_boq       匯入 BOQ
  boq_to_pr           BOQ 轉採購申請
```

**流程:** 讀 request.json → 解析 params → 建立 OdooClient → dispatch handler → 寫 response.json

### 4.2 bridge/odoo_client.py — Odoo API 客戶端

使用 `requests` + `HTTPBasicAuth`，呼叫 Odoo Swagger v2 端點。

| 方法 | HTTP | 端點 |
|------|------|------|
| `test_connection()` | GET | `/api/odoo-autocad/v2/test_connection` |
| `get_project(pr_no)` | PATCH | `/api/odoo-autocad/v2/get_project` |
| `get_products()` | PATCH | `/api/odoo-autocad/v2/get_product_v2` |
| `get_setup()` | PATCH | `/api/odoo-autocad/v2/get_setup_v2` |
| `get_colors(project_id)` | PATCH | `/api/odoo-autocad/v2/get_color_v2` |
| `import_to_boq(data)` | PATCH | `/api/odoo-autocad/v2/import2boq_v2` |
| `boq_to_pr(header_ids)` | PATCH | `/api/odoo-autocad/v2/boq2pr_v2` |

### 4.3 JSON 檔案交換格式

**Request:**
```json
{
  "action": "import_to_boq",
  "params": { "project_id": 123, "layouts": [...] }
}
```

**Response（成功）:**
```json
{
  "success": true,
  "data": { ... },
  "message": "Successfully imported to BOQ"
}
```

**Response（失敗）:**
```json
{
  "success": false,
  "error_code": "AUTH_FAILED",
  "message": "Authentication failed: invalid credentials"
}
```

**Error Codes:**
- `CONNECTION_ERROR` — 無法連線
- `TIMEOUT` — 逾時
- `AUTH_FAILED` — 認證失敗（HTTP 401）
- `HTTP_{status}` — 其他 HTTP 錯誤
- `MISSING_PARAM` — 缺少必要參數
- `UNKNOWN_ACTION` — 未知動作
- `FILE_NOT_FOUND` — 檔案不存在
- `INVALID_JSON` — JSON 解析錯誤
- `INTERNAL_ERROR` — 其他內部錯誤

---

## 5. 安全考量

- **認證資訊** 存放在 `config/bridge.ini`，不透過 JSON 檔案傳遞
- **暫存檔** 使用唯一檔名（含時間戳 + 流水號），用完即刪
- **Bridge.exe** 使用 HTTPS 與 Odoo 通訊
- **config/bridge.ini** 加入 `.gitignore`，提供 `.example` 範例
- **bridge.ini 密碼** 明碼存放（與現有 Python 應用一致，未來可加密）

---

## 6. 開發階段進度

### Phase 1: 基礎架構 — DONE
- [x] lisp/config.lsp — 路徑常數 + INI 讀寫
- [x] lisp/json_util.lsp — JSON 解析器
- [x] lisp/file_util.lsp — 檔案 I/O 工具
- [x] dcl/main_menu.dcl — 主選單對話框
- [x] dcl/result.dcl — 結果顯示對話框
- [x] dcl/config.dcl — 連線設定對話框

### Phase 2: Bridge 通訊 — DONE
- [x] lisp/odoo_bridge.lsp — 檔案交換 + 輪詢機制
- [x] bridge/odoo_bridge.py — CLI 入口 + action dispatch
- [x] bridge/odoo_client.py — Odoo Swagger API 客戶端
- [x] bridge/auth.py — BasicAuth 認證管理
- [x] bridge/config.py — INI 設定讀取
- [x] bridge/requirements.txt — Python 依賴

### Phase 3: AutoCAD 資料操作 — DONE
- [x] lisp/table_util.lsp — TABLE 遍歷 + 資料收集 + ID 回寫 + 清除
- [x] lisp/block_util.lsp — Block 屬性讀寫 + 搜尋
- [x] lisp/strip_mtext.lsp — MText 格式清除（RegExp + LT fallback）

### Phase 4: UI 邏輯 + 整合 — DONE
- [x] dcl/param_form.dcl — 參數表單（Product 關鍵字搜尋 + 6 下拉選單）
- [x] lisp/param_form.lsp — 載入 Odoo 資料 → 顯示表單 → 寫入屬性
- [x] lisp/main_menu.lsp — While 迴圈 + 7 個 action handler
- [x] lisp/main.lsp — 模組載入 + 8 個使用者指令

### Phase 5: 打包部署 — DONE
- [x] bridge/build.bat — PyInstaller 打包腳本
- [x] VLX 支援 — dcl:load 統一載入函數（VLX 內嵌 / 外部檔案自動判斷）
- [x] main.lsp — ob:vlx-mode-p 偵測，VLX 模式自動跳過 load

### 交付方式

| 格式 | AutoCAD Full | AutoCAD LT 2024+ | 程式碼保護 |
|------|:------------:|:-----------------:|:----------:|
| .lsp | 可 | 可 | 無（明文） |
| .fas | 可 | 可 | 編譯保護 |
| .vlx | 可 | 可 | 基本保護 |

#### 建置 VLX（需在 Full AutoCAD 中操作）

VLX 透過 AutoCAD 內建的 VLISP IDE 建置，步驟如下：

**步驟 1: 開啟 VLISP IDE**
```
在 AutoCAD 命令列輸入: VLIDE
→ 開啟 Visual LISP IDE 視窗
```

**步驟 2: 啟動 Make Application Wizard**
```
VLISP IDE 選單: File → Make Application → New Application Wizard...
```

**步驟 3: 設定應用程式屬性**
```
Application Name: OdooBridge
Application Location: 選擇輸出目錄（例如 autolisp/dist/）
Application Options:
  [x] Separate Namespace  ← 建議勾選，避免全域變數污染
```

**步驟 4: 加入 LISP 檔案（按順序）**
```
按 "Add..." 加入以下 10 個 .lsp 檔（順序重要，依相依性排列）：

  1. lisp/json_util.lsp
  2. lisp/file_util.lsp
  3. lisp/config.lsp
  4. lisp/strip_mtext.lsp
  5. lisp/block_util.lsp
  6. lisp/table_util.lsp
  7. lisp/odoo_bridge.lsp
  8. lisp/param_form.lsp
  9. lisp/main_menu.lsp
 10. lisp/main.lsp          ← 必須最後載入（進入點）
```

**步驟 5: 建置**
```
按 "Next" → "Finish"
→ VLISP 編譯所有 .lsp → 打包為 OdooBridge.vlx
→ 建置成功後會在 Console 顯示訊息
```

> **替代方式:** 也可在 AutoCAD 命令列直接輸入 `MAKELISPAPP`，
> 會開啟同樣的 Wizard 介面。

> **注意:** VLIDE / MAKELISPAPP 僅在 **Full AutoCAD** 中可用，
> AutoCAD LT 無此功能。但建置產出的 .vlx 可在 LT 2024+ 中執行。

> **DCL 嵌入:** Expert mode 的 "Resource Files" 頁面可將 .dcl 嵌入 VLX。
> 嵌入後交付時不需額外帶 dcl/ 目錄，`dcl:load` 會優先從 VLX 內部載入。
> 若未嵌入 DCL，`dcl:load` 會自動 fallback 到外部檔案路徑。

#### 交付包結構

```
dist/
├── OdooBridge.vlx           # 主程式（內嵌 .lsp + .dcl，編譯保護）
├── odoo_bridge.exe           # Python Bridge
└── config/bridge.ini         # 設定檔（使用者需修改連線資訊）
```

#### 安裝方式

**方式 A: APPLOAD 命令（推薦，永久載入）**
```
1. 將交付檔案複製到固定位置，例如:
   C:/OdooBridge/
     ├── OdooBridge.vlx
     ├── odoo_bridge.exe
     └── config/bridge.ini

2. 載入 VLX:
   → 命令列輸入: APPLOAD
   → 瀏覽選擇 C:/OdooBridge/OdooBridge.vlx
   → 按 "Load"

3. 設定每次啟動自動載入:
   → 在 APPLOAD 對話框中，找到 "Startup Suite" 區域
   → 按 "Contents..."
   → 按 "Add..." → 選擇 OdooBridge.vlx
   → 按 "Close"
   → 之後每次開啟 AutoCAD 都會自動載入
```

**方式 B: 命令列手動載入（臨時測試用）**
```
在 AutoCAD 命令列輸入:
  (load "C:/OdooBridge/OdooBridge.vlx")
```

**方式 C: acaddoc.lsp 自動載入**
```
在 AutoCAD 支援檔案搜尋路徑中建立或編輯 acaddoc.lsp，加入:
  (load "C:/OdooBridge/OdooBridge.vlx")
此檔案在每次開啟圖檔時自動執行。
```

**載入成功確認:**
```
命令列應顯示:
  [OB] ========================================
  [OB] Odoo-AutoCAD Integration (AutoLISP)
  [OB] Loading modules...
  [OB] VLX mode: json_util.lsp (embedded)
  [OB] VLX mode: file_util.lsp (embedded)
  ...
  [OB] All modules loaded successfully!
  [OB] Commands:
  [OB]   OB:MENU        - Main menu
  ...

輸入 OB:MENU 開啟主選單。
```

**常見問題排除:**
```
問題: OB:MENU 顯示 "Error: main_menu.dcl not found"
原因: VLX 建置時未嵌入 DCL（Simple mode 建置）
解法: 用 Expert mode 重新建置，在 Resource Files 頁加入 4 個 .dcl
     或將 dcl/ 目錄加入 OPTIONS → Files → Support File Search Path

問題: bridge.exe 找不到
原因: config.lsp 中的路徑設定不正確
解法: 編輯 config/bridge.ini，確認 [paths] bridge_exe 指向正確位置
```

### 待辦（Future）
- [ ] bridge/tests/ — Bridge 單元測試
- [ ] AutoCAD 內端對端測試
- [ ] 安裝腳本（自動設定搜尋路徑）
- [ ] 使用者操作手冊
