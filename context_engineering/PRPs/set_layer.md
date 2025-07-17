# PRP: set_layer

> **功能**: 設定 AutoCAD 的當前圖層  
> **類型**: MCP 工具  
> **優先級**: 中  
> **日期**: 2025年7月16日  

## 功能概述

### 目標
設定 AutoCAD 的當前圖層，如果圖層不存在則自動創建。這是第一階段圖層管理工具的重要組成部分。

### 使用場景
1. **圖層切換**: 工程師需要切換到不同的圖層進行繪圖
2. **圖層創建**: 創建新的圖層並設為當前圖層
3. **圖層管理**: 管理圖層的狀態和屬性
4. **AI 助手整合**: 透過自然語言指令設定圖層

### 成功標準
- [ ] 能夠成功設定當前圖層
- [ ] 支援圖層不存在時自動創建
- [ ] 支援圖層顏色和線型設定
- [ ] 提供清晰的操作回饋
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def set_layer(
    layer_name: str,
    color: int = 7,
    create_if_not_exist: bool = True
) -> Dict[str, Any]:
    """設定 AutoCAD 的當前圖層"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| layer_name | str | 是 | - | 圖層名稱 |
| color | int | 否 | 7 | 圖層顏色 (0-255) |
| create_if_not_exist | bool | 否 | True | 圖層不存在時是否自動創建 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "layer_name": "WALLS",
        "color": 7,
        "is_current": True,
        "created": False,
        "layer_info": {
            "name": "WALLS",
            "color": 7,
            "linetype": "Continuous",
            "lineweight": "Default",
            "on": True,
            "frozen": False,
            "locked": False
        },
        "changed_at": "2025-07-16T10:30:00"
    },
    "message": "成功設定圖層: WALLS",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "圖層名稱不能為空",
    "error_code": "INVALID_LAYER_NAME",
    "suggestion": "請提供有效的圖層名稱"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [ ] 需要 Odoo 連接（僅用於日誌記錄）
- [ ] 需要資料庫存取（僅用於配置）
- [x] 需要 win32com.client 和 pythoncom

### 核心邏輯
1. **參數驗證**: 檢查圖層名稱、顏色範圍
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **圖層檢查**: 驗證圖層是否存在
4. **圖層創建**: 如果圖層不存在且 create_if_not_exist 為 True，則創建圖層
5. **圖層設定**: 設定圖層顏色和屬性
6. **當前圖層設定**: 將圖層設為當前圖層
7. **結果回傳**: 返回圖層設定結果

### 錯誤處理
- **參數錯誤**: 圖層名稱無效、顏色範圍錯誤
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **圖層錯誤**: 圖層不存在且不允許創建、圖層創建失敗
- **設定錯誤**: 圖層設定失敗、權限問題

## 範例代碼

### 基本實現
```python
@mcp.tool()
def set_layer(
    layer_name: str,
    color: int = 7,
    create_if_not_exist: bool = True
) -> Dict[str, Any]:
    """設定 AutoCAD 的當前圖層"""
    logger.info(f"set_layer called with params: {locals()}")
    
    try:
        # 參數驗證
        if not isinstance(layer_name, str) or not layer_name.strip():
            return {
                "status": "error",
                "message": "圖層名稱不能為空",
                "error_code": "INVALID_LAYER_NAME"
            }
        
        # 驗證顏色範圍
        if not isinstance(color, int) or not (0 <= color <= 255):
            return {
                "status": "error",
                "message": "顏色必須是 0-255 之間的整數",
                "error_code": "INVALID_COLOR"
            }
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 設定圖層
        result = autocad_util.set_layer(
            layer_name=layer_name.strip(),
            color=color,
            create_if_not_exist=create_if_not_exist
        )
        
        return {
            "status": "success",
            "data": {
                "layer_name": result.get("layer_name"),
                "color": result.get("color"),
                "is_current": result.get("is_current"),
                "created": result.get("created"),
                "layer_info": result.get("layer_info"),
                "changed_at": datetime.now().isoformat()
            },
            "message": f"成功設定圖層: {result.get('layer_name')}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in set_layer: {e}")
        return {
            "status": "error",
            "message": f"設定圖層失敗: {str(e)}",
            "error_code": "LAYER_SETTING_ERROR"
        }
```

### 使用範例
```python
# 基本使用 - 設定現有圖層
result = set_layer(layer_name="WALLS")

# 設定圖層並指定顏色
result = set_layer(
    layer_name="DIMENSIONS",
    color=2  # 黃色
)

# 設定圖層但不自動創建
result = set_layer(
    layer_name="EXISTING_LAYER",
    create_if_not_exist=False
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_set_layer.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime

import mcp_server_fastmcp

class TestSetLayer:
    def test_set_layer_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "WALLS",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {
                "name": "WALLS",
                "color": 7,
                "linetype": "Continuous",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="WALLS")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer_name"] == "WALLS"
        assert result["data"]["color"] == 7
        assert result["data"]["is_current"] == True
        assert result["data"]["created"] == False
        assert "layer_info" in result["data"]
        assert "changed_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_set_layer_parameter_validation_empty_name(self):
        """參數驗證測試 - 空白圖層名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(layer_name="")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_set_layer_parameter_validation_whitespace_name(self):
        """參數驗證測試 - 只有空白的圖層名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(layer_name="   ")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_set_layer_parameter_validation_invalid_color_negative(self):
        """參數驗證測試 - 負數顏色"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST",
            color=-1
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COLOR"
        
    def test_set_layer_parameter_validation_invalid_color_high(self):
        """參數驗證測試 - 過高顏色值"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST",
            color=256
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COLOR"
        
    def test_set_layer_parameter_validation_invalid_color_type(self):
        """參數驗證測試 - 無效顏色類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST",
            color="red"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COLOR"
        
    def test_set_layer_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(layer_name="TEST")
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_set_layer_with_custom_color(self):
        """自定義顏色測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "DIMENSIONS",
            "color": 2,
            "is_current": True,
            "created": True,
            "layer_info": {
                "name": "DIMENSIONS",
                "color": 2,
                "linetype": "Continuous",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="DIMENSIONS",
            color=2
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer_name"] == "DIMENSIONS"
        assert result["data"]["color"] == 2
        assert result["data"]["created"] == True
        
    def test_set_layer_create_if_not_exist_false(self):
        """不自動創建圖層測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "EXISTING",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {
                "name": "EXISTING",
                "color": 7,
                "linetype": "Continuous",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="EXISTING",
            create_if_not_exist=False
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["created"] == False
        
    def test_set_layer_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TEST_LAYER",
            "color": 5,
            "is_current": True,
            "created": True,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST_LAYER",
            color=5,
            create_if_not_exist=True
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="TEST_LAYER",
            color=5,
            create_if_not_exist=True
        )
        
    def test_set_layer_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.side_effect = Exception("Layer error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.set_layer(layer_name="TEST")
        assert result["status"] == "error"
        assert result["error_code"] == "LAYER_SETTING_ERROR"
        
    def test_set_layer_layer_name_strip(self):
        """圖層名稱去除空白測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TRIMMED",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="  TRIMMED  ")
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="TRIMMED",
            color=7,
            create_if_not_exist=True
        )
        
    def test_set_layer_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'set_layer')
        assert callable(mcp_server_fastmcp.set_layer)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def set_layer(
    layer_name: str,
    color: int = 7,
    create_if_not_exist: bool = True
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    if not isinstance(layer_name, str) or not layer_name.strip():
        return {
            "status": "error",
            "message": "圖層名稱不能為空",
            "error_code": "INVALID_LAYER_NAME"
        }
    
    if not isinstance(color, int) or not (0 <= color <= 255):
        return {
            "status": "error",
            "message": "顏色必須是 0-255 之間的整數",
            "error_code": "INVALID_COLOR"
        }
    
    if not _autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    # 模擬結果
    return {
        "status": "success",
        "data": {
            "layer_name": layer_name.strip(),
            "color": color,
            "is_current": True,
            "created": False,
            "layer_info": {},
            "changed_at": datetime.now().isoformat()
        },
        "message": f"成功設定圖層: {layer_name.strip()}",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def set_layer(
    layer_name: str,
    color: int = 7,
    create_if_not_exist: bool = True
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_set_layer.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_set_layer.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_set_layer.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_set_layer.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_set_layer_integration.py
import pytest
from mcp_server_fastmcp import set_layer

class TestSetLayerIntegration:
    def test_set_layer_with_real_autocad(self):
        """與真實 AutoCAD 的整合測試"""
        # 注意：這個測試需要真實的 AutoCAD 連接
        # 假設 AutoCAD 已經啟動並連接
        pass
```

### 測試覆蓋率要求
- 單元測試覆蓋率: 95%+
- 整合測試覆蓋率: 90%+
- 所有錯誤路徑都要測試
- 所有參數組合都要測試

## 驗證標準

### 功能驗證
- [ ] 基本功能正常運作
- [ ] 參數驗證正確
- [ ] 錯誤處理完整
- [ ] 回傳值格式正確
- [ ] 支援所有指定的參數組合
- [ ] 與 AutoCAD 正常互動
- [ ] 圖層創建和設定正確

### 品質標準
- [ ] 代碼覆蓋率 ≥ 95%
- [ ] 所有測試通過
- [ ] 符合代碼風格指南
- [ ] 包含完整日誌記錄
- [ ] 遵循 TDD 原則
- [ ] 無靜態分析警告

### 整合標準
- [ ] 與現有 MCP 系統整合
- [ ] 與 AutoCAD 正常連接
- [ ] 與 Odoo 系統相容
- [ ] 符合 CLAUDE.md 規則
- [ ] 不影響現有功能
- [ ] 日誌格式一致

## 實施檢查清單

### 開發前
- [ ] 完成 PRP 審查
- [ ] 準備測試資料
- [ ] 確認依賴項目
- [ ] 設定 TDD 環境

### 開發中 (TDD 循環)
- [ ] 寫失敗測試 (Red)
- [ ] 最小實現 (Green)
- [ ] 重構改善 (Refactor)
- [ ] 遵循既有模式
- [ ] 實施錯誤處理
- [ ] 添加日誌記錄

### 開發後
- [ ] 執行所有測試
- [ ] 檢查代碼覆蓋率
- [ ] 進行代碼審查
- [ ] 更新文檔
- [ ] 整合測試
- [ ] 性能測試

## 相關資源

### 參考文件
- [AutoCAD MCP 整合方案](../../doc/AutoCAD_MCP_Integration_Plan.md)
- [CLAUDE.md](../../CLAUDE.md)
- [Context Engineering 分析](../../doc/Context_Engineering_Analysis.md)
- [TDD + Context Engineering 整合指南](../TDD_CONTEXT_ENGINEERING_GUIDE.md)

### 相關範例
- [MCP 工具範例](../examples/mcp_tool_template.py)
- [AutoCAD COM 範例](../examples/autocad_com_examples.py)
- [錯誤處理模式](../examples/error_handling_patterns.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_autocad.py` - AutoCAD 工具類別
- `tests/unit/test_mcp_*.py` - 現有測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。