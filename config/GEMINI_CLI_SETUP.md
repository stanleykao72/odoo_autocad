# Gemini CLI 配置指南

本指南說明如何配置 Gemini CLI 以連接到 AutoCAD-Odoo MCP 伺服器。

## 配置檔案選擇

我們提供三種不同的配置範本：

### 1. TCP Socket 配置 (推薦)
```json
// 檔案: gemini-cli-config-tcp.json
{
  "mcpServers": {
    "autocad-odoo-tcp": {
      "command": "C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe",
      "args": ["--mcp-server", "--mcp-port", "8000"],
      "transport": {
        "type": "tcp",
        "host": "localhost",
        "port": 8000
      }
    }
  }
}
```

**優點：**
- 跨平台相容性最佳
- 網路透明度
- 較容易除錯

### 2. Named Pipe 配置 (Windows 最佳化)
```json
// 檔案: gemini-cli-config-pipe.json
{
  "mcpServers": {
    "autocad-odoo-pipe": {
      "command": "C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe",
      "args": ["--mcp-server", "--mcp-pipe", "\\\\.\\pipe\\odoo_autocad_mcp"],
      "transport": {
        "type": "pipe",
        "name": "\\\\.\\pipe\\odoo_autocad_mcp"
      }
    }
  }
}
```

**優點：**
- Windows 上效能最佳
- 本地IPC，安全性較高
- 較低的延遲

### 3. 混合配置 (開發測試用)
```json
// 檔案: gemini-cli-config-hybrid.json
{
  "mcpServers": {
    "autocad-odoo": {
      "command": "C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe",
      "args": ["--mcp-server"],
      "transport": {
        "type": "tcp",
        "host": "localhost",
        "port": 8000
      }
    }
  }
}
```

## 安裝步驟

### 步驟 1: 確認應用程式安裝
確保 AutoCAD-Odoo 整合應用程式已安裝至：
```
C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe
```

### 步驟 2: 複製配置檔案
1. 選擇適合的配置範本 (建議使用 TCP 配置)
2. 複製配置檔案內容
3. 將內容貼到 Gemini CLI 的配置檔案中

### 步驟 3: 啟動 MCP 伺服器
有兩種方式啟動 MCP 伺服器：

#### 方式一：GUI 模式 (手動啟動)
```bash
# 啟動應用程式
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe"

# 在 GUI 中點擊 🚀 按鈕啟動 AI 助手服務
```

#### 方式二：純伺服器模式 (自動啟動)
```bash
# TCP 模式
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe" --mcp-server

# 自訂端口
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe" --mcp-server --mcp-port 8001

# Named Pipe 模式
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe" --mcp-server --mcp-pipe "\\\\.\\pipe\\custom_pipe"
```

### 步驟 4: 測試連接
啟動 Gemini CLI 並測試連接：

```bash
gemini-cli --config your_config.json
```

## 可用的 MCP 工具

連接成功後，你可以使用以下 AI 助手工具：

### AutoCAD 讀圖功能
- `scan_all_entities` - 掃描 AutoCAD 圖面中的所有實體
- `get_table_data` - 讀取圖面中的表格資料
- `extract_layout_info` - 提取圖面佈局資訊
- `query_entities_by_type` - 按實體類型查詢圖面元素

### Odoo 整合功能
- `get_products_by_query` - 按查詢條件搜尋 Odoo 產品
- `push_boq_to_project` - 推送 BOQ 資料到 Odoo 專案
- `validate_product_mapping` - 驗證產品代碼對應

## 範例對話

### 讀取 AutoCAD 圖面資訊
```
使用者: "請掃描目前 AutoCAD 圖面中的所有實體"
AI: 將調用 scan_all_entities 工具並返回詳細的實體清單
```

### 查詢 Odoo 產品
```
使用者: "搜尋所有包含 '管道' 關鍵字的產品"
AI: 將調用 get_products_by_query 工具搜尋相關產品
```

### 整合工作流程
```
使用者: "讀取圖面表格資料並推送到 Odoo 專案 123"
AI: 將依序調用 get_table_data 和 push_boq_to_project 工具
```

## 疑難排解

### 1. 連接失敗
- 確認 MCP 伺服器已啟動
- 檢查端口是否被其他應用程式使用
- 確認 Odoo 和 AutoCAD 連接設定正確

### 2. AutoCAD 工具無回應
- 確認 AutoCAD 應用程式已啟動
- 檢查 AutoCAD 文件是否已開啟
- 確認有足夠的權限存取 AutoCAD COM 介面

### 3. Odoo 工具錯誤
- 檢查 Odoo 伺服器連接設定
- 確認資料庫連接正常
- 驗證 API 權限設定

### 4. 日誌查看
在純伺服器模式下，日誌會輸出到控制台：
```bash
[LOG] MCP伺服器已啟動 - TCP:8000, Pipe:\\.\pipe\odoo_autocad_mcp
[LOG] TCP client connected: ('127.0.0.1', 12345)
[LOG] 收到MCP請求: {"jsonrpc": "2.0", "method": "tools/list", "id": 1}
```

## 支援聯絡

如有問題，請檢查：
1. 應用程式日誌檔案
2. Gemini CLI 錯誤訊息
3. AutoCAD 和 Odoo 連接狀態

---

**注意事項：**
- 確保 AutoCAD 和 Odoo 的連接設定正確
- MCP 伺服器需要在有效的 Odoo 和 AutoCAD 環境中運行
- 建議先在 GUI 模式中測試基本功能，再使用純伺服器模式