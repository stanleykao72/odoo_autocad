# 🚀 Gemini CLI + MCP 測試快速指南

## 📋 測試步驟 (5分鐘內完成)

### 1️⃣ 啟動應用程式
```bash
cd C:\odoo\autocad_source
python odoo.py
```
✅ **預期**: MCP Server 自動在端口 8083 啟動

### 2️⃣ 驗證 MCP Server
- 查看 GUI 日誌確認 "SSE 伺服器已啟動"
- 或開瀏覽器訪問: `http://localhost:8083` (應該有回應)

### 3️⃣ 設定 Gemini CLI
- **Server URL**: `http://localhost:8083`
- **Protocol**: SSE
- **確認連接成功**

### 4️⃣ 測試指令

#### 測試 1: 圓形繪製
```
畫一個半徑10的圓形
```

#### 測試 2: 線段繪製
```
從(0,0)到(100,100)畫一條線
```

### 5️⃣ 預期結果
✅ Gemini CLI 成功解析指令  
✅ MCP Server 回應包含執行詳情  
✅ 如果 AutoCAD 開啟 → 看到實際圖形  
✅ 處理時間 < 5秒  

---

## 🔧 快速故障排除

### ❌ Gemini CLI 無法連接
- 檢查: `netstat -ano | findstr :8083`
- 重啟: `python odoo.py`

### ❌ 指令無回應
- 確認 AutoCAD 已啟動
- 查看應用程式日誌

### ❌ 圖形未顯示
- ✅ 指令解析成功就算通過 (AutoCAD 可選)

---

## 🎯 成功標準

- [x] 2/2 指令成功解析
- [x] MCP Server 正常回應  
- [x] 處理時間 < 5秒
- [x] 無致命錯誤

**Task 1 里程碑: 早期端到端測試 ✅ 完成！**