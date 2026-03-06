# AutoLISP + Odoo Bridge

AutoCAD (LT / Full) 與 Odoo ERP 整合工具，採用純 AutoLISP + Python Bridge 架構。

## 架構

```
AutoCAD  ←→  AutoLISP  ←→  JSON 檔案  ←→  odoo_bridge.exe  ←→  Odoo API
```

- **AutoLISP** 負責 AutoCAD 內的 TABLE/Block 資料操作和 UI 對話框
- **Python Bridge** 負責與 Odoo Swagger API 的 HTTP 通訊
- 兩者透過 JSON 檔案交換資料，相容 AutoCAD LT 2024+

## 目錄結構

```
autolisp/
├── lisp/           AutoLISP 原始碼（LT 相容）
├── dcl/            DCL 對話框定義
├── bridge/         Python Bridge 原始碼
├── config/         設定檔
├── docs/           技術文件
├── legacy/         舊版檔案（參考用）
└── dist/           建置輸出
```

## 快速開始

### 1. 設定 Odoo 連線

複製 `config/bridge.ini.example` 為 `config/bridge.ini`，填入 Odoo 伺服器資訊。

### 2. 在 AutoCAD 中載入

```
(load "path/to/autolisp/lisp/main.lsp")
```

### 3. 可用指令

| 指令 | 功能 |
|------|------|
| `OB:CONNECT` | 測試 Odoo 連線 |
| `OB:SET-PROJECT` | 設定專案編號 |
| `OB:SET-PARAMS` | 參數選擇（產品/加工/顏色） |
| `OB:PUSH-BOQ` | 推送 BOQ 到 Odoo |
| `OB:CREATE-PR` | BOQ 轉採購申請 |
| `OB:CONFIG` | 設定連線資訊 |

## 開發

詳見 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
