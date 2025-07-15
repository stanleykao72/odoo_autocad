# AutoCAD-Odoo MCP 整合部署指南

本指南說明如何部署和配置 AutoCAD-Odoo MCP 整合系統。

## 🚀 快速開始

### 1. 建置可執行檔

```bash
# 清理並建置
python build_exe.py
```

建置完成後，可執行檔位於：`output/odoo-autocad-integration.exe`

### 2. 安裝到目標位置

複製到標準安裝路徑：
```
C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe
```

### 3. 測試基本功能

```bash
# 啟動GUI模式
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe"

# 測試純MCP伺服器模式
"C:/odoo/Odoo and AutoCAD Integration/odoo-autocad-integration.exe" --mcp-server
```

## 🔧 MCP 伺服器配置

### 支援的啟動模式

#### 1. GUI模式 (預設)
```bash
odoo-autocad-integration.exe
```
- 完整的圖形用戶介面
- 手動控制MCP伺服器啟動/停止
- 適合互動式使用

#### 2. GUI模式 + 自動啟動MCP
```bash
odoo-autocad-integration.exe --enable-mcp
```
- GUI介面 + 自動啟動MCP伺服器
- 最佳的混合模式

#### 3. 純MCP伺服器模式
```bash
odoo-autocad-integration.exe --mcp-server
```
- 無GUI，純命令列模式
- 適合服務化部署
- 支援自訂端口和管道

### 命令列參數

| 參數 | 描述 | 預設值 |
|------|------|--------|
| `--mcp-server` | 啟動純MCP伺服器模式 | GUI模式 |
| `--mcp-port` | TCP伺服器端口 | 8000 |
| `--mcp-pipe` | Named Pipe名稱 | `\\.\pipe\odoo_autocad_mcp` |
| `--enable-mcp` | GUI模式下自動啟動MCP | 手動控制 |

### 範例啟動指令

```bash
# 標準TCP伺服器
odoo-autocad-integration.exe --mcp-server

# 自訂端口
odoo-autocad-integration.exe --mcp-server --mcp-port 8001

# 自訂Named Pipe
odoo-autocad-integration.exe --mcp-server --mcp-pipe "\\.\pipe\custom_mcp"

# GUI + 自動MCP
odoo-autocad-integration.exe --enable-mcp
```

## 🌐 Gemini CLI 整合

### 配置檔案選擇

我們提供三種配置範本：

1. **TCP配置** (`config/gemini-cli-config-tcp.json`)
   - 跨平台相容性最佳
   - 預設端口 8000

2. **Named Pipe配置** (`config/gemini-cli-config-pipe.json`)
   - Windows最佳化效能
   - 本地IPC通訊

3. **混合配置** (`config/gemini-cli-config-hybrid.json`)
   - 開發測試用
   - 自動選擇最佳傳輸方式

### 快速配置

1. 選擇適合的配置範本
2. 修改路徑以符合實際安裝位置
3. 複製到Gemini CLI配置檔案

範例配置：
```json
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

## 🧪 測試和驗證

### 1. 連接測試

使用內建測試腳本：
```bash
python test_mcp_connection.py
```

### 2. 手動測試

使用netcat或telnet測試：
```bash
# 測試TCP連接
telnet localhost 8000

# 發送JSON-RPC請求
{"jsonrpc": "2.0", "method": "tools/list", "id": 1}
```

### 3. 功能驗證

測試各項MCP工具：

```bash
# 啟動MCP伺服器
odoo-autocad-integration.exe --mcp-server

# 在另一個終端測試連接
python test_mcp_connection.py
```

## 🛡️ 部署最佳實務

### 安全性考量

1. **本地連接**: 預設只監聽localhost，避免網路暴露
2. **權限控制**: 確保AutoCAD和Odoo的適當權限
3. **資料驗證**: 所有MCP請求都有輸入驗證

### 效能優化

1. **Named Pipe**: Windows環境優先使用Named Pipe
2. **連接池**: MCP伺服器支援多個並發連接
3. **資源管理**: 自動清理無效連接

### 監控和日誌

#### GUI模式
- 日誌顯示在底部視窗
- 即時狀態更新

#### 純伺服器模式
- 日誌輸出到控制台
- 支援重定向到檔案

```bash
# 將日誌重定向到檔案
odoo-autocad-integration.exe --mcp-server > mcp_server.log 2>&1
```

## 🔧 故障排除

### 常見問題

#### 1. 無法啟動MCP伺服器
```
錯誤: Address already in use
解決: 檢查端口是否被占用，使用 --mcp-port 指定其他端口
```

#### 2. AutoCAD連接失敗
```
錯誤: AutoCAD COM interface not available
解決: 確認AutoCAD已啟動且有適當權限
```

#### 3. Odoo連接問題
```
錯誤: Unable to connect to Odoo server
解決: 檢查config/下的Odoo配置檔案
```

### 診斷工具

1. **連接測試**: `python test_mcp_connection.py`
2. **端口檢查**: `netstat -an | findstr :8000`
3. **程序監控**: 工作管理員中查看`odoo-autocad-integration.exe`

### 日誌分析

關鍵日誌訊息：
```
[LOG] MCP伺服器已啟動 - TCP:8000, Pipe:\\.\pipe\odoo_autocad_mcp
[LOG] TCP client connected: ('127.0.0.1', 12345)
[LOG] 收到MCP請求: {"jsonrpc": "2.0", "method": "tools/list", "id": 1}
[LOG] MCP回應已發送
```

## 📚 使用範例

### 透過Gemini CLI使用

```
使用者: "請掃描AutoCAD圖面中的所有實體"
AI: 正在調用 scan_all_entities 工具...
    找到 15 個實體：
    - 5 條線段
    - 3 個圓形
    - 7 個文字標註

使用者: "搜尋包含'管道'的Odoo產品"
AI: 正在搜尋產品資料庫...
    找到 8 個相關產品：
    - 管道 100mm (P100)
    - 管道 150mm (P150)
    ...
```

### 批次處理工作流程

```
使用者: "讀取圖面表格資料並推送到Odoo專案123"
AI: 執行工作流程：
    1. 調用 get_table_data 讀取表格
    2. 驗證產品映射
    3. 調用 push_boq_to_project 推送資料
    完成！已成功推送 25 項BOQ資料
```

---

## 📞 技術支援

如遇問題，請收集以下資訊：
1. 錯誤訊息和日誌檔案
2. 系統環境 (Windows版本、AutoCAD版本)
3. Odoo連接配置
4. MCP伺服器啟動參數

詳細的故障排除指南請參考 `config/GEMINI_CLI_SETUP.md`