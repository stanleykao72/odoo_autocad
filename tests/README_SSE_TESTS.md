# SSE 測試指南

## 概述

本目錄包含了所有與 SSE (Server-Sent Events) 伺服器整合相關的測試。這些測試驗證了重新設計的 `MCPSSEManager` 類，該類使用直接整合方式而非子進程管理。

## 測試結構

```
tests/
├── unit/
│   └── test_direct_sse_manager.py      # 單元測試：直接 SSE 管理器
├── integration/
│   ├── test_simple_sse.py              # 整合測試：簡單 SSE 功能
│   └── test_gui_sse.py                 # 整合測試：GUI SSE 整合
├── tools/
│   └── verify_sse_integration.py       # 工具：完整整合驗證
└── run_sse_tests.py                    # 測試運行器
```

## 測試說明

### 單元測試

**test_direct_sse_manager.py**
- 測試直接整合的 `MCPSSEManager` 類
- 驗證伺服器啟動、停止、狀態查詢
- 測試直接工具調用功能
- 驗證 HTTP API 連接

### 整合測試

**test_simple_sse.py**
- 簡單的 SSE 伺服器功能測試
- 快速驗證基本連接性
- 適合持續整合環境

**test_gui_sse.py**
- 模擬 GUI 環境的 SSE 整合
- 測試狀態回調機制
- 驗證長期運行穩定性

### 工具

**verify_sse_integration.py**
- 完整的端到端整合驗證
- 測試 MCP JSON-RPC 協議相容性
- 提供 Gemini CLI 配置指導
- 適合手動驗證和故障排除

**diagnose_gui_sse.py**
- 診斷 GUI SSE 問題的專用工具
- 檢查伺服器內部狀態和線程情況
- 詳細的健康檢查和連接測試
- 適合調試 GUI 整合問題

**test_mcp_sse_compliance.py**
- 測試 MCP SSE 協議合規性
- 驗證是否符合 Gemini CLI 期望
- 包含完整的 MCP 握手流程測試
- 適合協議相容性驗證

**start_sse_for_gemini.py**
- 啟動 SSE 伺服器供 Gemini CLI 連接
- 保持運行直到手動停止
- 提供配置指導和狀態監控
- 適合外部 CLI 工具連接測試

## 運行測試

### 運行所有 SSE 測試
```bash
# 在專案根目錄
cd tests
python run_sse_tests.py
```

### 運行單個測試
```bash
# 單元測試
cd tests/unit
python test_direct_sse_manager.py

# 整合測試  
cd tests/integration
python test_simple_sse.py
python test_gui_sse.py

# 驗證工具
cd tests/tools
python verify_sse_integration.py
python diagnose_gui_sse.py
python test_mcp_sse_compliance.py

# 啟動 SSE 供外部連接
cd tests/tools
python start_sse_for_gemini.py
```

### 使用 pytest
```bash
# 運行所有 SSE 相關測試（如果有 pytest 標記）
pytest tests/ -k "sse" -v

# 運行特定類別
pytest tests/unit/ -v
pytest tests/integration/ -v
```

## 測試要求

### 環境依賴
- Python 3.10+
- FastAPI
- uvicorn
- requests
- 專案相關模組 (utility.util_mcp_sse_manager)

### 端口使用
- 測試會使用不同端口避免衝突：
  - `test_direct_sse_manager.py`: 8082
  - `test_simple_sse.py`: 8084
  - `test_gui_sse.py`: 8083
  - `verify_sse_integration.py`: 8083

### 注意事項
- 測試會啟動實際的 HTTP 伺服器
- 某些測試可能需要較長時間完成
- 確保測試端口未被其他程序佔用

## 故障排除

### 常見問題

**端口被佔用錯誤**
```
ERROR: [Errno 10048] error while attempting to bind on address
```
解決方案：
- 等待之前的測試完全結束
- 手動停止佔用端口的進程
- 使用不同的端口

**模組導入錯誤**
```
ModuleNotFoundError: No module named 'utility'
```
解決方案：
- 確保從正確的目錄運行測試
- 檢查 Python 路徑設定
- 使用提供的測試運行器

**編碼問題 (Windows)**
```
UnicodeEncodeError: 'cp950' codec can't encode
```
解決方案：
- 測試腳本已包含 UTF-8 編碼設定
- 在 PowerShell 中設定：`$env:PYTHONIOENCODING="utf-8"`

## 測試覆蓋範圍

這些測試涵蓋了：

✅ **核心功能**
- SSE 伺服器啟動/停止
- 直接 MCPSSEServer 整合
- 狀態管理和回調

✅ **API 相容性**
- MCP JSON-RPC 協議
- HTTP POST 端點
- SSE 串流端點

✅ **整合場景**
- GUI 環境模擬
- Gemini CLI 連接
- 錯誤處理和恢復

✅ **效能和穩定性**
- 長期運行測試
- 並發連接處理
- 記憶體洩漏檢查

## 未來改進

- [ ] 新增自動化 CI/CD 整合
- [ ] 實現效能基準測試
- [ ] 新增負載測試場景
- [ ] 擴展錯誤情境覆蓋
- [ ] 新增 mock 測試減少外部依賴