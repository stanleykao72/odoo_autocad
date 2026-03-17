# 安裝手冊 — Odoo and AutoCAD Integration v6.0

> 適用對象：終端使用者
> 版本：6.0.0.14
> 更新日期：2026-03-17

---

## 目錄

1. [系統需求](#1-系統需求)
2. [安裝流程總覽](#2-安裝流程總覽)
3. [Step 1：安裝 Windows 應用程式 (EXE)](#3-step-1安裝-windows-應用程式-exe)
4. [Step 2：設定 AutoCAD (VLX)](#4-step-2設定-autocad-vlx)
5. [Step 3：首次啟動與連線](#5-step-3首次啟動與連線)
6. [資料夾結構](#6-資料夾結構)
7. [COM 與 IPC 模式說明](#7-com-與-ipc-模式說明)
8. [常見問題排除](#8-常見問題排除)

---

## 1. 系統需求

| 項目 | 需求 |
|------|------|
| 作業系統 | Windows 10/11 (64-bit) |
| AutoCAD | Full AutoCAD 2014+ (COM 模式) 或 AutoCAD LT 2024+ (IPC 模式) |
| Odoo | v18+ 且已安裝 BOQ Import API 模組 |
| 磁碟空間 | 約 150 MB |

---

## 2. 安裝流程總覽

```
┌─────────────────────────────────────────────────┐
│  Step 1  安裝 Windows 應用程式 (EXE 安裝包)       │
│          → C:\odoo\Odoo and AutoCAD Integration  │
├─────────────────────────────────────────────────┤
│  Step 2  設定 AutoCAD (載入 VLX 檔案)             │
│          → McpDispatch.vlx + OdooAutoCAD.vlx     │
├─────────────────────────────────────────────────┤
│  Step 3  首次啟動、連線 Odoo 與 AutoCAD            │
└─────────────────────────────────────────────────┘
```

---

## 3. Step 1：安裝 Windows 應用程式 (EXE)

### 3.1 取得安裝檔

安裝檔位於：

```
installer\odoo-autocad-integration-6.0-setup.exe
```

### 3.2 執行安裝

1. 雙擊 `odoo-autocad-integration-6.0-setup.exe`
2. 若出現 Windows SmartScreen 警告，點選 **「其他資訊」→「仍要執行」**
3. 按照精靈指引完成安裝

> 預設安裝路徑：`C:\odoo\Odoo and AutoCAD Integration`

### 3.3 安裝完成後的檔案

```
C:\odoo\Odoo and AutoCAD Integration\
├── odoo-autocad-integration.exe    ← 主程式（已簽章）
├── config\                         ← 連線設定
│   ├── server.yaml
│   └── token.yaml
├── db\                             ← 本機快取資料庫
│   └── database.db
├── icon\                           ← 程式圖示
│   └── odoo_autocad.ico
└── fonts\                          ← 中文字型
```

---

## 4. Step 2：設定 AutoCAD (VLX)

系統需要在 AutoCAD 中載入兩個 VLX 檔案，以啟用資料提取功能。

### 4.1 複製 VLX 檔案

將以下兩個檔案複製到安裝目錄：

```
來源：autolisp\dist\McpDispatch.vlx
來源：autolisp\dist\OdooAutoCAD.vlx

目標：C:\odoo\Odoo and AutoCAD Integration\vlx\
```

完成後目錄結構：

```
C:\odoo\Odoo and AutoCAD Integration\
├── odoo-autocad-integration.exe
├── vlx\
│   ├── McpDispatch.vlx             ← IPC 通訊核心
│   └── OdooAutoCAD.vlx             ← Odoo 整合功能（8 個模組）
├── config\
│   ...
```

### 4.2 AutoCAD 設定搜尋路徑

1. 開啟 AutoCAD
2. 輸入指令 `OPTIONS` → 按 Enter
3. 切到 **「檔案 (Files)」** 頁籤
4. 展開 **「支援檔搜尋路徑 (Support File Search Path)」**
5. 點選 **「加入 (Add)」**，新增：

```
C:\odoo\Odoo and AutoCAD Integration\vlx
```

6. 按 **「確定 (OK)」** 儲存

### 4.3 載入 VLX（手動）

1. 在 AutoCAD 輸入指令 `APPLOAD` → 按 Enter
2. 瀏覽到 `C:\odoo\Odoo and AutoCAD Integration\vlx\`
3. **先選 `McpDispatch.vlx`** → 點選「載入 (Load)」
4. **再選 `OdooAutoCAD.vlx`** → 點選「載入 (Load)」

> **載入順序很重要！** McpDispatch 必須先載入，OdooAutoCAD 才能正常運作。

### 4.4 設定自動載入（建議）

為避免每次開啟 AutoCAD 都要手動載入，請加入 Startup Suite：

1. 在 `APPLOAD` 對話框中，找到右下角 **「啟動組合 (Startup Suite)」**
2. 點選 **「內容 (Contents)」**
3. 依序加入：
   - `C:\odoo\Odoo and AutoCAD Integration\vlx\McpDispatch.vlx`
   - `C:\odoo\Odoo and AutoCAD Integration\vlx\OdooAutoCAD.vlx`
4. 關閉對話框

> **注意：** Startup Suite 的載入順序依加入順序，請確保 McpDispatch 在前。

### 4.5 安全性設定

AutoCAD 可能會警告未簽章的 VLX 檔案。請依需求調整：

| SECURELOAD 值 | 行為 |
|---------------|------|
| 0 | 信任所有 VLX（最方便，開發環境建議） |
| 1 | 每次詢問是否載入（預設） |
| 2 | 只載入已簽章的 VLX |

設定方式：在 AutoCAD 指令列輸入：

```
SECURELOAD
輸入新值 <1>: 1
```

### 4.6 驗證載入成功

在 AutoCAD 指令列輸入：

```
OB:MCP-DISPATCH
```

若顯示 `"OB:MCP-DISPATCH registered"` 或無錯誤訊息，表示載入成功。

---

## 5. Step 3：首次啟動與連線

### 5.1 啟動應用程式

雙擊桌面捷徑或執行：

```
C:\odoo\Odoo and AutoCAD Integration\odoo-autocad-integration.exe
```

### 5.2 連線 Odoo

1. 點選左側 **「連接到 Odoo」**
2. 系統會使用 `config\` 中的設定連線
3. 右上角顯示 **「Odoo: 已連線」** 表示成功

### 5.3 連線 AutoCAD

1. 確認 AutoCAD 已開啟且有 .dwg 檔案
2. 選擇模式（左側切換按鈕）：
   - **COM** — Full AutoCAD 2014+
   - **IPC** — AutoCAD LT 2024+（需已載入 VLX）
3. 點選 **「連接到 AutoCAD」**
4. 右上角顯示 **「AutoCAD: 已連線」** 表示成功

### 5.4 基本操作流程

```
連線 Odoo → 連線 AutoCAD → 從 Odoo 獲取參數 → 推送到 BOQ → 轉移 BOQ 到 PR
```

---

## 6. 資料夾結構

### 完整安裝後的目錄

```
C:\odoo\Odoo and AutoCAD Integration\
│
├── odoo-autocad-integration.exe    ← 主程式（已簽章）
│
├── vlx\                            ← AutoCAD 擴充功能
│   ├── McpDispatch.vlx             ← IPC 核心（先載入）
│   └── OdooAutoCAD.vlx             ← Odoo 模組（後載入）
│
├── config\                         ← 連線設定
│   ├── server.yaml                 ← Odoo 伺服器位址
│   └── token.yaml                  ← API 認證金鑰
│
├── db\                             ← 本機快取
│   └── database.db                 ← SQLite 資料庫
│
├── icon\                           ← 程式圖示
│   └── odoo_autocad.ico
│
├── fonts\                          ← 中文字型
│   └── *.ttf
│
└── logs\                           ← 執行日誌（自動產生）
    └── odoo_autocad_YYYY-MM-DD.log
```

### AutoCAD 設定路徑

```
AutoCAD OPTIONS → Files → Support File Search Path:
  ✅ C:\odoo\Odoo and AutoCAD Integration\vlx

AutoCAD APPLOAD → Startup Suite:
  ✅ C:\odoo\Odoo and AutoCAD Integration\vlx\McpDispatch.vlx   (第 1 個)
  ✅ C:\odoo\Odoo and AutoCAD Integration\vlx\OdooAutoCAD.vlx   (第 2 個)
```

---

## 7. COM 與 IPC 模式說明

| 項目 | COM 模式 | IPC 模式 |
|------|----------|----------|
| 適用 AutoCAD | Full AutoCAD 2014+ | AutoCAD LT 2024+ / Full |
| 通訊方式 | Windows COM (pywin32) | File IPC (JSON 檔案交換) |
| 速度 | 快（毫秒級） | 中等（秒級） |
| VLX 需求 | 不需要 | 必須載入 McpDispatch.vlx + OdooAutoCAD.vlx |
| 線程安全 | 主線程執行（STA 安全） | 背景線程（ProgressDialog） |

### 如何選擇？

- 使用 **Full AutoCAD** → 選 **COM 模式**（預設，速度最快）
- 使用 **AutoCAD LT** → 必須選 **IPC 模式**（LT 不支援 COM）
- 使用 Full AutoCAD 但想用 IPC → 也可以，需載入 VLX

---

## 8. 常見問題排除

### Q: 連接 AutoCAD 時出現「請先連接 AutoCAD 並確保已獲取專案資料」

**原因**：AutoCAD 未開啟，或未開啟 .dwg 檔案
**解決**：先開啟 AutoCAD 並開啟圖檔，再點選「連接到 AutoCAD」

### Q: IPC 模式連線失敗

**原因**：VLX 未載入或載入順序錯誤
**解決**：
1. 在 AutoCAD 輸入 `APPLOAD`
2. 先載入 `McpDispatch.vlx`，再載入 `OdooAutoCAD.vlx`
3. 輸入 `OB:MCP-DISPATCH` 確認載入成功

### Q: AutoCAD 警告「無法載入未簽署的應用程式」

**解決**：輸入 `SECURELOAD`，設定為 `1`（提示後可選擇載入）或 `0`（信任全部）

### Q: 推送到 BOQ 失敗

**可能原因**：
- Odoo 未連線 → 先點「連接到 Odoo」
- 圖檔中沒有 TABLE 或 Block → 確認圖檔有正確的表格與屬性區塊
- Odoo 資料衝突 → 查看系統日誌 (`logs\` 目錄) 中的錯誤訊息

### Q: 日誌檔在哪裡？

```
C:\odoo\Odoo and AutoCAD Integration\logs\odoo_autocad_YYYY-MM-DD.log
```

開啟日誌檔可查看詳細的操作記錄與錯誤訊息。
