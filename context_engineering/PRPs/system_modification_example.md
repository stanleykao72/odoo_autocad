# 系統修改 PRP 範例：新增 create_new_drawing 工具

## 修改概述
在既有的 `mcp_server_fastmcp.py` 中新增 `create_new_drawing` MCP 工具，以支援創建新的 AutoCAD 圖面。

## 目標系統
- **檔案**: `mcp_server_fastmcp.py`
- **當前狀態**: 已有 7 個 MCP 工具
- **修改類型**: 功能擴展（新增第 8 個工具）

## 整合要求

### 遵循既有模式
```python
# 必須遵循的既有模式
@mcp.tool()
def new_tool_name(parameters) -> Dict[str, Any]:
    """工具描述"""
    logger.info(f"Tool called with: {locals()}")
    
    try:
        # 實現邏輯
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error(f"Error: {e}")
        return {"status": "error", "message": str(e)}
```

### 使用既有資源
- **AutoCAD 連接**: 使用 `self.autocad_util` 實例
- **錯誤處理**: 使用既有的 try-catch 模式
- **日誌系統**: 使用既有的 logger
- **回傳格式**: 遵循既有的 Dict[str, Any] 格式

## 具體修改內容

### 在 mcp_server_fastmcp.py 中新增
```python
@mcp.tool()
def create_new_drawing(
    template_path: str = "",
    drawing_name: str = "新圖面",
    units: str = "公制"
) -> Dict[str, Any]:
    """創建新的 AutoCAD 圖面
    
    Args:
        template_path: 模板檔案路徑（可選）
        drawing_name: 圖面名稱
        units: 單位系統（公制/英制）
    
    Returns:
        Dict containing success status and drawing info
    """
    logger.info(f"create_new_drawing called with: {locals()}")
    
    try:
        # 檢查 AutoCAD 連接
        if not self.autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 創建新圖面的邏輯
        result = self.autocad_util.create_new_drawing(
            template_path=template_path,
            drawing_name=drawing_name,
            units=units
        )
        
        return {
            "status": "success",
            "message": f"成功創建圖面: {drawing_name}",
            "data": {
                "drawing_name": drawing_name,
                "units": units,
                "template_used": template_path or "預設模板"
            }
        }
        
    except Exception as e:
        logger.error(f"Error in create_new_drawing: {e}")
        return {
            "status": "error",
            "message": f"創建圖面失敗: {str(e)}",
            "error_code": "DRAWING_CREATION_ERROR"
        }
```

### 可能需要的輔助修改
如果 `UtilAutoCAD` 類別沒有 `create_new_drawing` 方法，則需要在 `utility/util_autocad.py` 中新增。

## 測試要求

### 既有功能測試
- [ ] 確保所有既有的 7 個 MCP 工具仍正常運作
- [ ] AutoCAD 連接狀態不受影響
- [ ] Odoo 整合功能正常

### 新功能測試
- [ ] `create_new_drawing` 工具正常運作
- [ ] 錯誤處理正確
- [ ] 日誌記錄完整

## 向後相容性
- 不修改既有工具的簽名
- 不改變既有的回傳格式
- 保持既有的錯誤處理機制

## 驗收標準
- [ ] 新工具通過所有測試
- [ ] 既有功能無回歸問題
- [ ] 符合既有代碼風格
- [ ] 整合測試通過

---

**重點**: 這是系統修改型 PRP，重點在於與既有系統的整合和相容性。