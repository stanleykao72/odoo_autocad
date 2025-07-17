# PRP: draw_line

> **功能**: 在 AutoCAD 中繪製直線  
> **類型**: MCP 工具  
> **優先級**: 中  
> **日期**: 2025年7月16日  

## 功能概述

### 目標
在 AutoCAD 中繪製直線，支援指定起點和終點座標，並可設定圖層屬性。這是第一階段基礎繪圖工具的重要組成部分。

### 使用場景
1. **基礎繪圖**: 工程師需要在圖面上繪製直線
2. **座標繪圖**: 根據精確座標繪製直線
3. **圖層管理**: 將直線繪製到指定圖層
4. **AI 助手整合**: 透過自然語言指令繪製直線

### 成功標準
- [ ] 能夠根據座標成功繪製直線
- [ ] 支援設定起點和終點座標
- [ ] 支援設定直線的圖層
- [ ] 提供清晰的操作回饋
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def draw_line(
    start_point: list,
    end_point: list,
    layer: str = "0"
) -> Dict[str, Any]:
    """在 AutoCAD 中繪製直線"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| start_point | list | 是 | - | 起點座標 [x, y, z] |
| end_point | list | 是 | - | 終點座標 [x, y, z] |
| layer | str | 否 | "0" | 圖層名稱 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "line_id": "AcDbLine:1234567890",
        "start_point": [0.0, 0.0, 0.0],
        "end_point": [100.0, 100.0, 0.0],
        "layer": "0",
        "length": 141.42,
        "angle": 45.0,
        "created_at": "2025-07-16T10:30:00"
    },
    "message": "成功繪製直線",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "座標格式不正確",
    "error_code": "INVALID_COORDINATES",
    "suggestion": "座標必須是包含3個數值的列表，如 [x, y, z]"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [ ] 需要 Odoo 連接（僅用於日誌記錄）
- [ ] 需要資料庫存取（僅用於配置）
- [x] 需要 win32com.client 和 pythoncom

### 核心邏輯
1. **參數驗證**: 檢查座標格式、座標數值有效性、圖層名稱
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **圖層檢查**: 驗證圖層是否存在，如不存在則創建
4. **直線繪製**: 使用 AutoCAD COM API 繪製直線
5. **屬性設定**: 設定直線的圖層屬性
6. **計算屬性**: 計算直線長度和角度
7. **結果回傳**: 返回繪製的直線資訊

### 錯誤處理
- **參數錯誤**: 座標格式不正確、座標數值無效、圖層名稱無效
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **圖層錯誤**: 圖層名稱無效、圖層創建失敗
- **繪製錯誤**: 直線繪製失敗、屬性設定失敗

## 範例代碼

### 基本實現
```python
@mcp.tool()
def draw_line(
    start_point: list,
    end_point: list,
    layer: str = "0"
) -> Dict[str, Any]:
    """在 AutoCAD 中繪製直線"""
    logger.info(f"draw_line called with params: {locals()}")
    
    try:
        # 參數驗證
        if not isinstance(start_point, list) or len(start_point) != 3:
            return {
                "status": "error",
                "message": "起點座標必須是包含3個數值的列表",
                "error_code": "INVALID_START_POINT"
            }
        
        if not isinstance(end_point, list) or len(end_point) != 3:
            return {
                "status": "error",
                "message": "終點座標必須是包含3個數值的列表",
                "error_code": "INVALID_END_POINT"
            }
        
        # 檢查座標數值
        try:
            start_coords = [float(x) for x in start_point]
            end_coords = [float(x) for x in end_point]
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "座標必須是數值",
                "error_code": "INVALID_COORDINATE_VALUES"
            }
        
        # 檢查是否為重複點
        if start_coords == end_coords:
            return {
                "status": "error",
                "message": "起點和終點不能相同",
                "error_code": "IDENTICAL_POINTS"
            }
        
        # 檢查圖層名稱
        if not isinstance(layer, str) or not layer.strip():
            return {
                "status": "error",
                "message": "圖層名稱必須是非空字串",
                "error_code": "INVALID_LAYER_NAME"
            }
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 繪製直線
        result = autocad_util.draw_line(
            start_point=start_coords,
            end_point=end_coords,
            layer=layer.strip()
        )
        
        # 計算直線屬性
        import math
        dx = end_coords[0] - start_coords[0]
        dy = end_coords[1] - start_coords[1]
        dz = end_coords[2] - start_coords[2]
        
        length = math.sqrt(dx*dx + dy*dy + dz*dz)
        angle = math.degrees(math.atan2(dy, dx))
        
        return {
            "status": "success",
            "data": {
                "line_id": result.get("line_id"),
                "start_point": start_coords,
                "end_point": end_coords,
                "layer": layer.strip(),
                "length": round(length, 2),
                "angle": round(angle, 2),
                "created_at": datetime.now().isoformat()
            },
            "message": f"成功繪製直線，長度: {round(length, 2)}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in draw_line: {e}")
        return {
            "status": "error",
            "message": f"繪製直線失敗: {str(e)}",
            "error_code": "LINE_DRAWING_ERROR"
        }
```

### 使用範例
```python
# 基本使用 - 繪製簡單直線
result = draw_line(
    start_point=[0, 0, 0],
    end_point=[100, 100, 0]
)

# 指定圖層
result = draw_line(
    start_point=[0, 0, 0],
    end_point=[200, 0, 0],
    layer="LINE_LAYER"
)

# 3D 直線
result = draw_line(
    start_point=[0, 0, 0],
    end_point=[100, 100, 50],
    layer="3D_LINES"
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_draw_line.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
import math

import mcp_server_fastmcp

class TestDrawLine:
    def test_draw_line_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.return_value = {
            "line_id": "AcDbLine:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["start_point"] == [0.0, 0.0, 0.0]
        assert result["data"]["end_point"] == [100.0, 100.0, 0.0]
        assert result["data"]["layer"] == "0"
        assert "length" in result["data"]
        assert "angle" in result["data"]
        
    def test_draw_line_parameter_validation_invalid_start_point(self):
        """參數驗證測試 - 無效起點"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0],  # 只有2個元素
            end_point=[100, 100, 0]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_START_POINT"
        
    def test_draw_line_parameter_validation_invalid_end_point(self):
        """參數驗證測試 - 無效終點"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point="invalid"  # 不是列表
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_END_POINT"
        
    def test_draw_line_parameter_validation_non_numeric_coordinates(self):
        """參數驗證測試 - 非數值座標"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, "z"],  # 非數值
            end_point=[100, 100, 0]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COORDINATE_VALUES"
        
    def test_draw_line_parameter_validation_identical_points(self):
        """參數驗證測試 - 相同起終點"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[0, 0, 0]  # 相同點
        )
        assert result["status"] == "error"
        assert result["error_code"] == "IDENTICAL_POINTS"
        
    def test_draw_line_parameter_validation_invalid_layer(self):
        """參數驗證測試 - 無效圖層"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0],
            layer=""  # 空字串
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_draw_line_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_draw_line_length_calculation(self):
        """長度計算測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.return_value = {
            "line_id": "AcDbLine:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[3, 4, 0]  # 3-4-5 直角三角形
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["length"] == 5.0
        
    def test_draw_line_angle_calculation(self):
        """角度計算測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.return_value = {
            "line_id": "AcDbLine:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[1, 1, 0]  # 45度角
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["angle"] == 45.0
        
    def test_draw_line_with_custom_layer(self):
        """自定義圖層測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.return_value = {
            "line_id": "AcDbLine:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0],
            layer="CUSTOM_LAYER"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer"] == "CUSTOM_LAYER"
        
    def test_draw_line_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.return_value = {
            "line_id": "AcDbLine:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[10, 20, 30],
            end_point=[40, 50, 60],
            layer="TEST_LAYER"
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.draw_line.assert_called_once_with(
            start_point=[10.0, 20.0, 30.0],
            end_point=[40.0, 50.0, 60.0],
            layer="TEST_LAYER"
        )
        
    def test_draw_line_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.side_effect = Exception("AutoCAD error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "LINE_DRAWING_ERROR"
        
    def test_draw_line_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'draw_line')
        assert callable(mcp_server_fastmcp.draw_line)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def draw_line(
    start_point: list,
    end_point: list,
    layer: str = "0"
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    import math
    
    # 基本驗證
    if not isinstance(start_point, list) or len(start_point) != 3:
        return {
            "status": "error",
            "message": "起點座標必須是包含3個數值的列表",
            "error_code": "INVALID_START_POINT"
        }
    
    if not isinstance(end_point, list) or len(end_point) != 3:
        return {
            "status": "error",
            "message": "終點座標必須是包含3個數值的列表",
            "error_code": "INVALID_END_POINT"
        }
    
    try:
        start_coords = [float(x) for x in start_point]
        end_coords = [float(x) for x in end_point]
    except (ValueError, TypeError):
        return {
            "status": "error",
            "message": "座標必須是數值",
            "error_code": "INVALID_COORDINATE_VALUES"
        }
    
    if start_coords == end_coords:
        return {
            "status": "error",
            "message": "起點和終點不能相同",
            "error_code": "IDENTICAL_POINTS"
        }
    
    if not isinstance(layer, str) or not layer.strip():
        return {
            "status": "error",
            "message": "圖層名稱必須是非空字串",
            "error_code": "INVALID_LAYER_NAME"
        }
    
    if not _autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    # 計算基本屬性
    dx = end_coords[0] - start_coords[0]
    dy = end_coords[1] - start_coords[1]
    dz = end_coords[2] - start_coords[2]
    
    length = math.sqrt(dx*dx + dy*dy + dz*dz)
    angle = math.degrees(math.atan2(dy, dx))
    
    return {
        "status": "success",
        "data": {
            "line_id": "AcDbLine:1234567890",
            "start_point": start_coords,
            "end_point": end_coords,
            "layer": layer.strip(),
            "length": round(length, 2),
            "angle": round(angle, 2),
            "created_at": datetime.now().isoformat()
        },
        "message": f"成功繪製直線，長度: {round(length, 2)}",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def draw_line(
    start_point: list,
    end_point: list,
    layer: str = "0"
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_draw_line.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_draw_line.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_draw_line.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_draw_line.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_draw_line_integration.py
import pytest
from mcp_server_fastmcp import draw_line

class TestDrawLineIntegration:
    def test_draw_line_with_real_autocad(self):
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
- [ ] 長度和角度計算正確

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