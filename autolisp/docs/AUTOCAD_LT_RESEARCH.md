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

### 方案 C：獨立 Python Bridge .exe（採用方案）

```
AutoLISP → 寫 request.json
         → startapp odoo_bridge.exe action request.json response.json
         → 輪詢等待 response.json
         → 讀取並解析結果
```

- 優點：邏輯清晰、好維護、可處理認證/重試/錯誤
- 缺點：需部署一個額外 .exe（可用 PyInstaller 打包）

### 方案比較

| 項目 | A (curl) | B (MSHTA) | **C (Python Bridge)** |
|------|----------|-----------|----------------------|
| LT 支援 | 可 | 可能被攔截 | **可** |
| 維護性 | 中 | 低 | **高** |
| 錯誤處理 | 差 | 差 | **好** |
| 安全性 | 中 | 低 | **高** |
| 部署複雜度 | 低 | 低 | 中 |

---

## 5. 採用方案 C 的理由

1. **與現有架構差異最小** — 只是把 `vlax-get-or-create-object "Python.ComServer"` 替換為 `startapp odoo_bridge.exe`，通訊從 COM 改為 JSON 檔案
2. **Python 技術棧一致** — 現有團隊已熟悉 Python + Odoo API
3. **可用 PyInstaller 打包** — 使用者不需安裝 Python 環境
4. **完整的錯誤處理** — 認證、重試、超時、HTTP 狀態碼都在 Bridge 內處理
5. **LT + 完整版通用** — 同一套 AutoLISP 腳本可同時在 LT 和完整版運行

---

## 參考來源

- [CAD Forum - LISP limitations in AutoCAD LT](https://www.cadforum.cz/en/limitations-of-the-lisp-language-autolisp-visuallisp-autocad-lt-tip13683)
- [CAD Forum - vlax-create-object workaround (Tip 14007)](https://www.cadforum.cz/en/workaround-for-the-missing-vlax-create-object-in-autocad-lt-tip14007)
- [Autodesk - AutoLISP support for AutoCAD LT](https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/Are-VBA-SCR-and-LISP-working-in-AutoCAD-LT.html)
- [Autodesk Blog - AutoCAD LT and AutoLISP](https://www.autodesk.com/blogs/autocad/autocad-lt-2024-autolisp/)
- [AutoCAD DevBlog - Entitlement API with LISP (XMLHTTP example)](https://adndevblog.typepad.com/autocad/2022/05/using-entitlement-api-with-lisp.html)
- [Autodesk Forum - REST POST API with LISP](https://forums.autodesk.com/t5/visual-lisp-autolisp-and-general/lisp-function-with-rest-post-api-is-it-doable/td-p/13282194)
