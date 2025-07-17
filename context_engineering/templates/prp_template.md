# PRP: [功能名稱]

> **功能**: [簡短功能描述]  
> **類型**: MCP 工具  
> **優先級**: [高/中/低]  
> **日期**: [建立日期]  

## 功能概述

### 目標
[詳細描述功能的目標和用途]

### 使用場景
[列出主要使用場景]

### 成功標準
[明確定義功能成功的標準]

## 技術規格

### 函數簽名
```python
@mcp.tool()
def [function_name]([parameters]) -> Dict[str, Any]:
    """[功能描述]"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| param1 | str | 是 | - | 參數描述 |
| param2 | int | 否 | 10 | 參數描述 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "result": [實際結果],
    "message": "操作成功",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "錯誤描述",
    "error_code": "ERROR_001",
    "suggestion": "建議解決方案"
}
```

## 實施要求

### 依賴項目
- [ ] 需要 AutoCAD 連接
- [ ] 需要 Odoo 連接
- [ ] 需要資料庫存取
- [ ] 其他依賴

### 核心邏輯
1. **參數驗證**: [描述驗證邏輯]
2. **主要處理**: [描述核心邏輯]
3. **結果回傳**: [描述回傳邏輯]

### 錯誤處理
- **參數錯誤**: [處理方式]
- **連接錯誤**: [處理方式]
- **執行錯誤**: [處理方式]

## 範例代碼

### 基本實現
```python
@mcp.tool()
def [function_name]([parameters]) -> Dict[str, Any]:
    """[功能描述]"""
    logger.info(f"[function_name] called with params: {locals()}")
    
    try:
        # 參數驗證
        if not param1:
            return {
                "status": "error",
                "message": "param1 is required",
                "error_code": "PARAM_ERROR"
            }
        
        # 取得必要的工具實例
        autocad_util = get_autocad_util()
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD utility not available",
                "error_code": "AUTOCAD_ERROR"
            }
        
        # 核心邏輯實現
        result = perform_main_operation(param1, param2)
        
        return {
            "status": "success",
            "result": result,
            "message": "操作成功完成",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in [function_name]: {e}")
        return {
            "status": "error",
            "message": str(e),
            "error_code": "EXECUTION_ERROR"
        }
```

### 使用範例
```python
# 基本使用
result = [function_name](param1="value1")

# 進階使用
result = [function_name](param1="value1", param2=20)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_[function_name].py
import pytest
from unittest.mock import patch, Mock

class Test[FunctionName]:
    def test_[function_name]_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        expected = {"status": "success", "data": {...}}
        
        # Act
        result = [function_name](valid_params)
        
        # Assert
        assert result == expected
        
    def test_[function_name]_parameter_validation(self):
        """參數驗證測試 - 應該失敗"""
        result = [function_name](invalid_params)
        assert result["status"] == "error"
        
    def test_[function_name]_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試 - 應該失敗"""
        with patch('mcp_server_fastmcp.autocad_util') as mock_autocad:
            mock_autocad.side_effect = Exception("Connection failed")
            result = [function_name](valid_params)
            assert result["status"] == "error"
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def [function_name]([parameters]) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    return {"status": "success", "data": {}}
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，保持測試通過
@mcp.tool()
def [function_name]([parameters]) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    logger.info(f"[function_name] called with: {locals()}")
    
    try:
        # 完整的實現邏輯
        # 參數驗證
        # AutoCAD 操作
        # 結果處理
        return {"status": "success", "data": result}
    except Exception as e:
        logger.error(f"Error in [function_name]: {e}")
        return {"status": "error", "message": str(e)}
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_[function_name].py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_[function_name].py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_[function_name].py -v  # 持續通過
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_[function_name]_integration.py
def test_[function_name]_with_real_autocad():
    """與真實 AutoCAD 的整合測試"""
    # 需要真實的 AutoCAD 連接
    pass
```

### 測試覆蓋率要求
- 單元測試覆蓋率: 95%+
- 整合測試覆蓋率: 90%+
- 所有錯誤路徑都要測試

## 驗證標準

### 功能驗證
- [ ] 基本功能正常運作
- [ ] 參數驗證正確
- [ ] 錯誤處理完整
- [ ] 回傳值格式正確

### 品質標準
- [ ] 代碼覆蓋率 ≥ 95%
- [ ] 所有測試通過
- [ ] 符合代碼風格指南
- [ ] 包含完整日誌記錄

### 整合標準
- [ ] 與現有 MCP 系統整合
- [ ] 與 AutoCAD 正常連接
- [ ] 與 Odoo 系統相容
- [ ] 符合 CLAUDE.md 規則

## 實施檢查清單

### 開發前
- [ ] 完成 PRP 審查
- [ ] 準備測試資料
- [ ] 確認依賴項目

### 開發中
- [ ] 遵循 TDD 方法
- [ ] 實施錯誤處理
- [ ] 添加日誌記錄

### 開發後
- [ ] 執行所有測試
- [ ] 進行代碼審查
- [ ] 更新文檔

## 相關資源

### 參考文件
- [AutoCAD MCP 整合方案](../../doc/AutoCAD_MCP_Integration_Plan.md)
- [CLAUDE.md](../../CLAUDE.md)
- [Context Engineering 分析](../../doc/Context_Engineering_Analysis.md)

### 相關範例
- [MCP 工具範例](../examples/mcp_tool_template.py)
- [AutoCAD COM 範例](../examples/autocad_com_examples.py)
- [錯誤處理模式](../examples/error_handling_patterns.py)

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。