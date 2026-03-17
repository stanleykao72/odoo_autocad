# Gemini CLI 配置指南 v5.0

> **版本**: 5.0  
> **發布日期**: 2024年7月15日  
> **MCP工具數量**: 7個（4個AutoCAD + 3個Odoo）

本指南說明如何配置 Gemini CLI 以連接到 AutoCAD-Odoo MCP 伺服器。

## 配置檔案選擇

我們提供三種不同的配置範本：

### 1. TCP Socket 配置 (需注意事項)
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

**注意：** 此配置可能導致 Gemini CLI 掛起。建議使用方式二或三。

**優點：**
- 跨平台相容性最佳
- 網路透明度
- 較容易除錯

**已知問題：**
- 應用程式的日誌輸出可能干擾 MCP 協議
- 可能導致 Gemini CLI 啟動時掛起

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

### 3. 官方 FastMCP 伺服器 (推薦)
```json
// 檔案: gemini-cli-config-fastmcp.json
{
  "mcpServers": {
    "autocad-odoo": {
      "command": "python",
      "args": ["C:/odoo/autocad_source/mcp_server_fastmcp.py"],
      "env": {
        "PYTHONUNBUFFERED": "1",
        "PYTHONDONTWRITEBYTECODE": "1"
      }
    }
  }
}
```

**優點：**
- 使用官方 FastMCP SDK，相容性最佳
- 支援協議版本 "2024-11-05"
- 自動處理 JSON-RPC 協議細節
- 包含測試工具和診斷功能
- 無 GUI 依賴，啟動速度快
- 完善的錯誤處理和協議驗證

### 3.1. SSE 模式伺服器 (實驗性)
```json
// 檔案: gemini-cli-config-sse.json
{
  "mcpServers": {
    "autocad-odoo-sse": {
      "transport": {
        "type": "sse",
        "url": "http://localhost:8080/sse"
      },
      "description": "AutoCAD-Odoo Integration with SSE transport",
      "startCommand": {
        "command": "python",
        "args": ["C:/odoo/autocad_source/mcp_server_sse.py", "8080"],
        "env": {
          "PYTHONUNBUFFERED": "1",
          "PYTHONDONTWRITEBYTECODE": "1"
        }
      }
    }
  }
}
```

**優點：**
- 支援 Server-Sent Events 即時串流
- HTTP 端點易於除錯和監控
- 支援 CORS 跨域請求
- 可同時處理多個並發連接
- 適合需要即時更新的應用場景

**注意：**
- SSE 模式在 MCP 2024-11-05 中已被棄用
- 需要手動啟動 HTTP 伺服器
- 某些 MCP 客戶端可能不支援 SSE 傳輸

### 4. GUI 模式配置
```json
// 檔案: gemini-cli-config-gui.json
{
  "mcpServers": {
    "autocad-odoo": {
      "command": "C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe",
      "args": ["--enable-mcp"],
      "env": {
        "PYTHONUNBUFFERED": "1"
      }
    }
  }
}
```

**優點：**
- 同時啟動 GUI 和 MCP 服務
- 避免純伺服器模式的掛起問題
- 可在 GUI 中監控狀態

### 5. 分離式伺服器配置
如果需要使用純伺服器模式，建議分兩步執行：

1. 手動啟動 MCP 伺服器：
```bash
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe" --mcp-server --mcp-port 8000
```

2. 在 Gemini CLI 中連接已啟動的伺服器：
```json
{
  "mcpServers": {
    "autocad-odoo-client": {
      "transport": {
        "type": "tcp",
        "host": "localhost",
        "port": 8000,
        "connect_only": true
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

### 步驟 2: 自動配置 (推薦)
使用自動化配置腳本，基於 [Gemini CLI Issue #3470](https://github.com/google-gemini/gemini-cli/issues/3470) 的建議：

```bash
# 設定推薦配置
python setup_mcp_config.py setup-recommended

# 或設定所有配置選項
python setup_mcp_config.py setup-all

# 列出當前配置
python setup_mcp_config.py list

# 驗證配置
python setup_mcp_config.py validate
```

### 步驟 3: 手動配置 (進階用戶)
1. 選擇適合的配置範本 (建議使用 FastMCP 配置)
2. 複製配置檔案內容
3. 將內容貼到 Gemini CLI 的配置檔案中

### 步驟 4: 啟動 MCP 伺服器
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

### 步驟 5: 測試連接
啟動 Gemini CLI 並測試連接：

```bash
gemini-cli --config your_config.json
```

## 配置範例

我們提供了多種配置範例，可在 `mcp_config_examples/` 目錄中找到：

### 基本配置
- `stdio_basic.json` - 基本 stdio 傳輸
- `sse_basic.json` - 基本 SSE 傳輸
- `http_basic.json` - 基本 HTTP 傳輸

### 進階配置
- `stdio_with_custom_env.json` - 包含自定義環境變數的 stdio
- `sse_with_headers.json` - 包含自定義標頭和認證的 SSE
- `http_with_auth.json` - 包含 Bearer 認證的 HTTP

### 特殊配置
- `multiple_servers.json` - 多伺服器配置
- `development_config.json` - 開發環境配置
- `production_config.json` - 生產環境配置

### 生成配置範例
```bash
# 生成所有配置範例
python mcp_config_examples.py

# 檢查生成的範例
ls mcp_config_examples/
```

## 可用的 MCP 工具

連接成功後，你可以使用以下 AI 助手工具：

### 基本測試工具 (改良版伺服器)
- `test_connection` - 測試 MCP 連接和伺服器狀態
- `get_server_info` - 獲取伺服器資訊和功能
- `check_autocad_status` - 檢查 AutoCAD 連接狀態
- `check_odoo_status` - 檢查 Odoo 連接狀態

### AutoCAD 讀圖功能 (完整版伺服器)
- `scan_all_entities` - 掃描 AutoCAD 圖面中的所有實體
- `get_table_data` - 讀取圖面中的表格資料
- `extract_layout_info` - 提取圖面佈局資訊
- `query_entities_by_type` - 按實體類型查詢圖面元素

### Odoo 整合功能 (完整版伺服器)
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

### 1. Gemini CLI 掛起問題
**症狀：** 使用 TCP 配置時，Gemini CLI 啟動後無回應

**原因：** 
- 應用程式的日誌輸出干擾 MCP 協議
- 純伺服器模式 (`--mcp-server`) 的 stdout/stderr 輸出與 MCP 協議衝突

**解決方案：**
- **推薦：** 使用官方 FastMCP 伺服器 (`mcp_server_fastmcp.py`)
- 使用 GUI 模式配置 (`--enable-mcp`)
- 或使用 Named Pipe 配置
- 或手動啟動伺服器後使用分離式配置

### 2. 連接失敗
- 確認 MCP 伺服器已啟動
- 檢查端口是否被其他應用程式使用
- 確認 Odoo 和 AutoCAD 連接設定正確

### 3. AutoCAD 工具無回應
- 確認 AutoCAD 應用程式已啟動
- 檢查 AutoCAD 文件是否已開啟
- 確認有足夠的權限存取 AutoCAD COM 介面

### 4. Odoo 工具錯誤
- 檢查 Odoo 伺服器連接設定
- 確認資料庫連接正常
- 驗證 API 權限設定

### 5. 日誌查看
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