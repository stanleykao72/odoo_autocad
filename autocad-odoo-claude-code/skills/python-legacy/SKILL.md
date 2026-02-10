# Python 舊系統分析 Skill

---
name: python-legacy
description: 研究分析 Python 版本的 Odoo-AutoCAD 整合程式碼，輔助 C# 遷移
trigger-keywords:
  - Python
  - 舊程式
  - 遷移
  - migration
  - pywin32
  - Bravado
  - Tkinter
  - util_autocad
  - util_odoo
  - 舊系統
  - legacy
allowed-tools:
  - Read
  - Grep
  - Glob
  - Task
---

## 概述

本 Skill 用於研究和分析 Python 版本的 Odoo-AutoCAD 整合應用程式，為 C# WPF 遷移提供參考。

## Python 原始碼結構

```
(專案根目錄)
├── odoo.py                          # 主程式入口
├── forms/
│   ├── form_main.py                 # 主視窗 (Tkinter)
│   ├── form_main_modern.py          # 現代化主視窗 (CustomTkinter)
│   └── form_autocad_param.py        # AutoCAD 參數對話框
├── utility/
│   ├── util_autocad.py              # AutoCAD COM 介面 ★
│   ├── util_odoo.py                 # Odoo API 客戶端 ★
│   ├── util_push_to_boq.py          # BOQ 推送邏輯 ★
│   ├── util_transfer_boq_to_pr.py   # BOQ→PR 轉換 ★
│   ├── util_com_server.py           # COM 伺服器
│   ├── util_gui_proxy.py            # GUI 代理 (v5.1)
│   └── util_mcp_sse_manager.py      # MCP SSE 管理
├── models/
│   └── server.py                    # SQLAlchemy 模型
├── config/
│   └── *.yaml                       # 環境配置
├── mcp_server_fastmcp.py            # MCP 伺服器
└── tests/                           # 測試目錄
```

## 關鍵模組分析指引

### util_autocad.py — AutoCAD COM 操作
**關注點**:
- `win32com.client.Dispatch` 呼叫方式
- 參數提取的欄位列表
- 圖層遍歷邏輯
- 錯誤處理模式

**C# 遷移對照**:
| Python | C# |
|--------|-----|
| `win32com.client.Dispatch("AutoCAD.Application")` | `Marshal.GetActiveObject("AutoCAD.Application")` |
| `acad.ActiveDocument` | `_acadApp.ActiveDocument` (dynamic) |
| 字典動態存取 | `dynamic` 或 Interop Assembly |

### util_odoo.py — Odoo API 整合
**關注點**:
- 認證流程（session-based）
- `search_read` 呼叫格式
- 資料快取策略
- 錯誤重試機制

**C# 遷移對照**:
| Python | C# |
|--------|-----|
| `requests` / `bravado` | `HttpClient` / `IHttpClientFactory` |
| `yaml.load()` | `IConfiguration` + `appsettings.json` |
| `SQLAlchemy` | `Entity Framework Core` |
| `threading` | `async/await` + `Task` |

### util_push_to_boq.py — BOQ 處理
**關注點**:
- BOQ 資料結構
- 驗證規則
- Odoo API 呼叫序列
- 錯誤回滾邏輯

### util_transfer_boq_to_pr.py — PR 生成
**關注點**:
- PR 資料映射
- Odoo 購買流程 API
- 批次處理邏輯

## 分析工作流

1. **識別功能**: 確認要遷移的 Python 功能
2. **讀取原始碼**: 使用 `Read` 工具讀取 Python 檔案
3. **提取介面**: 識別公開方法和資料結構
4. **對照設計**: 設計 C# 對應的介面和實作
5. **記錄差異**: 紀錄 Python/C# 的行為差異

## 常用搜尋指令

```
# 搜尋 AutoCAD COM 呼叫
Grep: pattern="Dispatch|GetActiveObject|ActiveDocument" path="utility/"

# 搜尋 Odoo API 端點
Grep: pattern="search_read|create|write|unlink" path="utility/util_odoo.py"

# 搜尋資料模型定義
Grep: pattern="class.*Model|Column\(" path="models/"

# 搜尋配置載入
Grep: pattern="yaml.load|config\[" path="utility/"
```

## 檢查清單

- [ ] 已讀取並理解 Python 原始碼
- [ ] 已識別核心業務邏輯
- [ ] 已記錄 Python → C# 對應關係
- [ ] 已標註無法直接遷移的部分
- [ ] 已確認測試覆蓋計畫
