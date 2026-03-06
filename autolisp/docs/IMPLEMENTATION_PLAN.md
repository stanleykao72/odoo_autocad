# AutoLISP DCL Main UI + Python Bridge — 實作計畫

> 版本: 1.0
> 日期: 2026-03-06
> 狀態: Phase 1-5 全部實作完成

---

## 1. 背景與目標

將現有 Python Tkinter GUI（`forms/form_main.py`, `form_autocad_param.py`）的功能，以純 AutoLISP + DCL 重新實作，相容 AutoCAD LT 2024+。Odoo 通訊改用 Python Bridge .exe + JSON 檔案交換（方案 C），取代原本的 Python COM Server。

### 核心需求

- **LT 相容**: 所有 `.lsp` 檔不含 `vlax-create-object` / `vlax-get-or-create-object`（例外：`strip_mtext.lsp` 使用 `vl-catch-all-apply` 包裹 + LT fallback）
- **DCL 對話框**: 取代 Tkinter GUI
- **Python Bridge**: 取代 Python COM Server 進行 Odoo 通訊
- **JSON 檔案交換**: `startapp` + 輪詢取代 COM method invoke

---

## 2. DCL 輪詢限制與解法

### 問題

`start_dialog()` 會阻塞執行，期間無法呼叫 `(command)`、無法寫入主控台、無計時器機制。

### 解法：While 迴圈 + done_dialog(N)

每個按鈕用 `done_dialog(N)` 關閉對話框 → 在對話框外執行工作（含 Bridge 輪詢）→ 顯示結果 → while 迴圈重新開啟主選單。

```lisp
(defun menu:show ()
  (setq keep-open T)
  (while keep-open
    ;; 載入 + 顯示對話框
    (setq dch (load_dialog dcl-path))
    (new_dialog "ob_main_menu" dch)
    (set_tile "status_text" (if *ob:odoo-connected* "Odoo: Connected" "Odoo: Not Connected"))
    (action_tile "btn_connect"    "(done_dialog 1)")
    (action_tile "btn_config"     "(done_dialog 2)")
    (action_tile "btn_set_params" "(done_dialog 3)")
    (action_tile "btn_push_boq"   "(done_dialog 4)")
    (action_tile "btn_create_pr"  "(done_dialog 5)")
    (action_tile "cancel"         "(done_dialog 0)")

    (setq action (start_dialog))   ;; 阻塞，等使用者按按鈕
    (unload_dialog dch)

    ;; === 對話框已關閉，可自由執行任何操作 ===
    (cond
      ((= action 0) (setq keep-open nil))   ;; Close
      ((= action 1) (menu:do-connect))      ;; 呼叫 bridge → 輪詢 → alert 結果
      ((= action 2) (menu:do-config))       ;; 開啟設定對話框
      ((= action 3) (menu:do-set-params))   ;; 載入 Odoo 資料 → 開參數表單
      ((= action 4) (menu:do-push-boq))     ;; 收集 TABLE → bridge → 回寫 ID
      ((= action 5) (menu:do-create-pr))    ;; bridge → 顯示結果
    )
    ;; while 迴圈自動重新顯示主選單（除非 action=0）
  )
)
```

### Bridge 輪詢在對話框外執行

完全沒有 DCL 限制問題：

```lisp
(defun bridge:wait-for-response (resp-file timeout / elapsed)
  (setq elapsed 0)
  (while (and (not (findfile resp-file)) (< elapsed timeout))
    (command "_.delay" 500)
    (setq elapsed (+ elapsed 500))
  )
  (if (findfile resp-file)
    (bridge:read-json resp-file)
    '(("success" . nil) ("error_code" . "TIMEOUT") ("message" . "Bridge timeout"))
  )
)
```

---

## 3. DCL 主選單設計

```
+------------------------------------------+
|  Odoo-AutoCAD Integration                |
+------------------------------------------+
|  Status: [Odoo: Not Connected]           |
+------------------------------------------+
| Connection                               |
|  [Test Odoo Connection]                  |
|  [Connection Settings...]                |
+------------------------------------------+
| Main Functions                           |
|  [Set Parameters from Odoo]              |
|  [Push to BOQ]                           |
|  [Transfer BOQ to PR]                    |
+------------------------------------------+
| Tools                                    |
|  [Clear Current Layout Table IDs]        |
|  [Clear ALL Layout Table IDs]            |
+------------------------------------------+
|              [Close]                     |
+------------------------------------------+
```

每個按鈕 → `(done_dialog N)` → while 迴圈 dispatch → 工作完成 → 重新顯示。

---

## 4. 實作檔案清單

### Phase 1: 基礎框架 — DONE

| # | 檔案 | 內容 | 狀態 |
|---|------|------|------|
| 1 | `lisp/config.lsp` | 路徑常數、INI 讀寫（純 AutoLISP `open`/`read-line`） | DONE |
| 2 | `lisp/json_util.lsp` | 從 `legacy/json_util.lsp` 複製，不修改 | DONE |
| 3 | `lisp/file_util.lsp` | `file:write-string`, `file:read-string`, `file:unique-name`, `file:delete` | DONE |
| 4 | `dcl/main_menu.dcl` | 主選單（7 按鈕 + 狀態列） | DONE |
| 5 | `dcl/result.dcl` | 結果顯示（title + 5-line message + OK） | DONE |
| 6 | `dcl/config.dcl` | 連線設定（URL, DB, username, password） | DONE |

### Phase 2: Bridge 通訊 — DONE

| # | 檔案 | 內容 | 狀態 |
|---|------|------|------|
| 7 | `lisp/odoo_bridge.lsp` | `bridge:call`（寫 JSON → startapp → 輪詢 → 讀回應）+ 7 個 wrapper | DONE |
| 8 | `bridge/odoo_bridge.py` | CLI 入口：`odoo_bridge.py <action> <req.json> <resp.json>` | DONE |
| 9 | `bridge/odoo_client.py` | Odoo Swagger API client（requests + BasicAuth） | DONE |
| 10 | `bridge/auth.py` | 認證管理（BasicAuth） | DONE |
| 11 | `bridge/config.py` | INI 讀取（configparser） | DONE |
| 12 | `bridge/requirements.txt` | `requests`, `configparser` | DONE |

### Phase 3: AutoCAD 資料操作 — DONE

| # | 檔案 | 來源 | 狀態 |
|---|------|------|------|
| 13 | `lisp/table_util.lsp` | 重構自 `legacy/transfer_to_odoo.lsp`（TABLE 遍歷、ID 回寫、清除） | DONE |
| 14 | `lisp/block_util.lsp` | 重構自 `legacy/contract_product.lsp`（Block 屬性讀寫） | DONE |
| 15 | `lisp/strip_mtext.lsp` | 從 legacy 取 `LM:UnFormat`，加 LT fallback（純字串替換） | DONE |

### Phase 4: UI 邏輯 + 整合 — DONE

| # | 檔案 | 內容 | 狀態 |
|---|------|------|------|
| 16 | `dcl/param_form.dcl` | 參數表單（Product 關鍵字搜尋 + 6 下拉選單 + UOM/color_no 自動填入） | DONE |
| 17 | `lisp/param_form.lsp` | 載入 Odoo 資料 → 顯示表單 → 寫入 Block 屬性 | DONE |
| 18 | `lisp/main_menu.lsp` | while 迴圈 + action dispatch + 所有 handler | DONE |
| 19 | `lisp/main.lsp` | 載入所有模組、註冊 `c:OB:MENU` 等 8 個指令 | DONE |

### Phase 5: 打包部署 — DONE

| # | 檔案 | 內容 | 狀態 |
|---|------|------|------|
| 20 | `bridge/build.bat` | PyInstaller 打包 → `dist/odoo_bridge.exe` | DONE |
| 21 | `lisp/file_util.lsp` | 新增 `dcl:load` — VLX/外部檔案雙模式 DCL 載入 | DONE |
| 22 | `lisp/main.lsp` | 新增 `ob:vlx-mode-p` — VLX 模式跳過 module load | DONE |

---

## 5. LT 相容注意事項

| 問題 | 解法 |
|------|------|
| `vlax-create-object` 不可用 | 所有外部通訊由 Bridge 處理，AutoLISP 端不使用 |
| `LM:UnFormat` 用 VBScript.RegExp | 加 LT fallback：`vl-catch-all-apply` 包裹 + 純 `vl-string-subst` 鏈式替換 MText 格式碼 |
| `read_csv.lsp` 的 ADODB | 不再使用 CSV，改由 Bridge 從 Odoo API 取資料 |
| `startapp` 非同步 | 輪詢在 DCL 外執行，用 `(command "_.delay" 500)` + `findfile` |
| DCL `start_dialog` 阻塞 | While 迴圈 + `done_dialog(N)` 關閉 → 外部工作 → 重新開啟 |

---

## 6. 使用者指令對照

| 舊指令 (COM) | 新指令 | 功能 |
|-------------|--------|------|
| `L0` | `OB:CONNECT` | 測試 Odoo 連線 |
| `L1` | `OB:SET-PARAMS` | 從 Odoo 載入選項 → 參數選擇表單 → 寫入 Block 屬性 |
| `L3` | `OB:PUSH-BOQ` | 收集所有 Layout TABLE → 推送 BOQ → 回寫 ID |
| `L4` | `OB:CREATE-PR` | 收集 header_ids → BOQ 轉 PR |
| (contract_product) | `OB:SET-PARAMS` | 同上（合併功能） |
| (new) | `OB:CONFIG` | 設定 Odoo 連線資訊（儲存至 bridge.ini） |
| (new) | `OB:MENU` | 主選單（上述所有功能） |
| `L7` | `OB:CLEAR-IDS` | 清除目前 Layout 的 TABLE ID |
| (new) | `OB:CLEAR-ALL-IDS` | 清除所有 Layout 的 TABLE ID |

---

## 7. 驗證方式

1. **Phase 1 驗證**: `(load "main.lsp")` → `OB:MENU` → 主選單正確顯示，按鈕可點擊
2. **Phase 2 驗證**: 點 "Test Connection" → Bridge 啟動 → 輪詢成功 → result dialog 顯示結果
3. **Phase 3 驗證**: 手動呼叫 `(table:get-all-layouts-data)` → 正確回傳 TABLE 資料
4. **Phase 4 驗證**: 完整流程：Set Params → Push BOQ → 回寫 ID → Create PR
5. **Phase 5 驗證**:
   - `bridge\build.bat` → 產出 `dist/odoo_bridge.exe` → 可獨立執行
   - Full AutoCAD: `VLIDE` → Make Application Wizard → 產出 `OdooBridge.vlx`
   - `(load "OdooBridge.vlx")` → 模組載入訊息顯示 "VLX mode" → `OB:MENU` 正常開啟
6. **LT 驗證**: 所有 `.lsp` 檔不含 bare `vlax-create-object` / `vlax-get-or-create-object`（strip_mtext.lsp 使用 catch-all + fallback）
7. **VLX + LT 驗證**: 在 AutoCAD LT 2024+ 中 `APPLOAD` → 載入 `OdooBridge.vlx` → `OB:MENU` 正常運作

### Bridge CLI 驗證（已通過）

```bash
$ cd autolisp/bridge
$ python odoo_bridge.py --help
usage: odoo_bridge.py [-h] [--config CONFIG] action request_file response_file

$ echo '{"action":"test_connection","params":{}}' > req.json
$ python odoo_bridge.py test_connection req.json resp.json
$ cat resp.json
{"success": false, "error_code": "CONNECTION_FAILED", "message": "Cannot connect to https://odoo.example.com"}
```

---

## 8. 關鍵參考檔案

| 檔案 | 用途 |
|------|------|
| `autolisp/legacy/transfer_to_odoo.lsp` | TABLE 操作邏輯藍本 |
| `autolisp/legacy/contract_product.lsp` | Block 屬性 + DCL 表單模式 |
| `autolisp/legacy/json_util.lsp` | JSON 解析器（直接沿用） |
| `autolisp/docs/ARCHITECTURE.md` | 方案 C 架構設計 |
| `autolisp/docs/AUTOCAD_LT_RESEARCH.md` | LT 相容性研究報告 |
