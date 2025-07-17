# PRP: draw_circle

> **功能**: 在 AutoCAD 中繪製圓形  
> **類型**: MCP 工具  
> **優先級**: 中  
> **日期**: 2025年7月16日  

## 功能概述

### 目標
在 AutoCAD 中繪製圓形，支援指定圓心座標和半徑，並可設定圖層屬性。這是第一階段基礎繪圖工具的重要組成部分。

### 使用場景
1. **基礎繪圖**: 工程師需要在圖面上繪製圓形
2. **座標繪圖**: 根據精確座標和半徑繪製圓形
3. **圖層管理**: 將圓形繪製到指定圖層
4. **AI 助手整合**: 透過自然語言指令繪製圓形

### 成功標準
- [ ] 能夠根據圓心座標和半徑成功繪製圓形
- [ ] 支援設定圓心座標和半徑
- [ ] 支援設定圓形的圖層
- [ ] 提供清晰的操作回饋
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def draw_circle(
    center_point: list,
    radius: float,
    layer: str = "0"
) -> Dict[str, Any]:
    """在 AutoCAD 中繪製圓形"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| center_point | list | 是 | - | 圓心座標 [x, y, z] |
| radius | float | 是 | - | 圓形半徑 |
| layer | str | 否 | "0" | 圖層名稱 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "circle_id": "AcDbCircle:1234567890",
        "center_point": [0.0, 0.0, 0.0],
        "radius": 50.0,
        "layer": "0",
        "area": 7853.98,
        "circumference": 314.16,
        "created_at": "2025-07-16T10:30:00"
    },
    "message": "成功繪製圓形",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "半徑必須為正數",
    "error_code": "INVALID_RADIUS",
    "suggestion": "半徑必須是大於0的數值"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [ ] 需要 Odoo 連接（僅用於日誌記錄）
- [ ] 需要資料庫存取（僅用於配置）
- [x] 需要 win32com.client 和 pythoncom

### 核心邏輯
1. **參數驗證**: 檢查圓心座標格式、半徑有效性、圖層名稱
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **圖層檢查**: 驗證圖層是否存在，如不存在則創建
4. **圓形繪製**: 使用 AutoCAD COM API 繪製圓形
5. **屬性設定**: 設定圓形的圖層屬性
6. **計算屬性**: 計算圓形面積和周長
7. **結果回傳**: 返回繪製的圓形資訊

### 錯誤處理
- **參數錯誤**: 圓心座標格式不正確、半徑無效、圖層名稱無效
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **圖層錯誤**: 圖層名稱無效、圖層創建失敗
- **繪製錯誤**: 圓形繪製失敗、屬性設定失敗

## 範例代碼

### 基本實現
```python
@mcp.tool()
def draw_circle(
    center_point: list,
    radius: float,
    layer: str = "0"
) -> Dict[str, Any]:
    """在 AutoCAD 中繪製圓形"""
    logger.info(f"draw_circle called with params: {locals()}")
    
    import math
    
    try:
        # 參數驗證
        if not isinstance(center_point, list) or len(center_point) != 3:
            return {
                "status": "error",
                "message": "圓心座標必須是包含3個數值的列表",
                "error_code": "INVALID_CENTER_POINT"
            }
        
        # 檢查座標數值
        try:
            center_coords = [float(x) for x in center_point]
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "圓心座標必須是數值",
                "error_code": "INVALID_CENTER_VALUES"
            }
        
        # 檢查半徑
        try:
            radius_value = float(radius)
            if radius_value <= 0:
                return {
                    "status": "error",
                    "message": "半徑必須為正數",
                    "error_code": "INVALID_RADIUS"
                }
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "半徑必須是數值",
                "error_code": "INVALID_RADIUS_TYPE"
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
        
        # 繪製圓形
        result = autocad_util.draw_circle(
            center_point=center_coords,
            radius=radius_value,
            layer=layer.strip()
        )
        
        # 計算圓形屬性
        area = math.pi * radius_value * radius_value
        circumference = 2 * math.pi * radius_value
        
        return {
            "status": "success",
            "data": {
                "circle_id": result.get("circle_id"),
                "center_point": center_coords,
                "radius": radius_value,
                "layer": layer.strip(),
                "area": round(area, 2),
                "circumference": round(circumference, 2),
                "created_at": datetime.now().isoformat()
            },
            "message": f"成功繪製圓形，半徑: {radius_value}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in draw_circle: {e}")
        return {
            "status": "error",
            "message": f"繪製圓形失敗: {str(e)}",
            "error_code": "CIRCLE_DRAWING_ERROR"
        }
```

### 使用範例
```python
# 基本使用 - 繪製簡單圓形
result = draw_circle(
    center_point=[0, 0, 0],
    radius=50
)

# 指定圖層
result = draw_circle(
    center_point=[100, 100, 0],
    radius=25,
    layer="CIRCLE_LAYER"
)

# 3D 圓形
result = draw_circle(
    center_point=[0, 0, 50],
    radius=30,
    layer="3D_CIRCLES"
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_draw_circle.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
import math

import mcp_server_fastmcp

class TestDrawCircle:
    def test_draw_circle_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.return_value = {
            "circle_id": "AcDbCircle:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["center_point"] == [0.0, 0.0, 0.0]
        assert result["data"]["radius"] == 50.0
        assert result["data"]["layer"] == "0"
        assert "area" in result["data"]
        assert "circumference" in result["data"]
        assert "circle_id" in result["data"]
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_draw_circle_parameter_validation_invalid_center_point(self):
        """參數驗證測試 - 無效圓心"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0],  # 只有2個元素
            radius=50
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_CENTER_POINT"
        
    def test_draw_circle_parameter_validation_invalid_radius_negative(self):
        """參數驗證測試 - 負半徑"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=-10  # 負數半徑
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_RADIUS"
        
    def test_draw_circle_parameter_validation_invalid_radius_zero(self):
        """參數驗證測試 - 零半徑"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=0  # 零半徑
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_RADIUS"
        
    def test_draw_circle_parameter_validation_invalid_radius_type(self):
        """參數驗證測試 - 無效半徑類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius="invalid"  # 非數值
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_RADIUS_TYPE"
        
    def test_draw_circle_parameter_validation_non_numeric_center_coordinates(self):
        """參數驗證測試 - 非數值圓心座標"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, "z"],  # 非數值
            radius=50
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_CENTER_VALUES"
        
    def test_draw_circle_parameter_validation_invalid_layer(self):
        """參數驗證測試 - 無效圖層"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50,
            layer=""  # 空字串
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_draw_circle_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50
        )
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_draw_circle_area_calculation(self):
        """面積計算測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.return_value = {
            "circle_id": "AcDbCircle:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=10  # 半徑10，面積應為100π
        )
        
        # Assert
        assert result["status"] == "success"
        expected_area = math.pi * 10 * 10
        assert abs(result["data"]["area"] - expected_area) < 0.01
        
    def test_draw_circle_circumference_calculation(self):
        """周長計算測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.return_value = {
            "circle_id": "AcDbCircle:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=10  # 半徑10，周長應為20π
        )
        
        # Assert
        assert result["status"] == "success"
        expected_circumference = 2 * math.pi * 10
        assert abs(result["data"]["circumference"] - expected_circumference) < 0.01
        
    def test_draw_circle_with_custom_layer(self):
        """自定義圖層測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.return_value = {
            "circle_id": "AcDbCircle:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50,
            layer="CUSTOM_LAYER"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer"] == "CUSTOM_LAYER"
        
    def test_draw_circle_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.return_value = {
            "circle_id": "AcDbCircle:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[10, 20, 30],
            radius=25,
            layer="TEST_LAYER"
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.draw_circle.assert_called_once_with(
            center_point=[10.0, 20.0, 30.0],
            radius=25.0,
            layer="TEST_LAYER"
        )
        
    def test_draw_circle_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.side_effect = Exception("AutoCAD error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50
        )
        assert result["status"] == "error"
        assert result["error_code"] == "CIRCLE_DRAWING_ERROR"
        
    def test_draw_circle_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'draw_circle')
        assert callable(mcp_server_fastmcp.draw_circle)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def draw_circle(
    center_point: list,
    radius: float,
    layer: str = "0"
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    import math
    
    # 基本驗證
    if not isinstance(center_point, list) or len(center_point) != 3:
        return {
            "status": "error",
            "message": "圓心座標必須是包含3個數值的列表",
            "error_code": "INVALID_CENTER_POINT"
        }
    
    try:
        center_coords = [float(x) for x in center_point]
    except (ValueError, TypeError):
        return {
            "status": "error",
            "message": "圓心座標必須是數值",
            "error_code": "INVALID_CENTER_VALUES"
        }
    
    try:
        radius_value = float(radius)
        if radius_value <= 0:
            return {
                "status": "error",
                "message": "半徑必須為正數",
                "error_code": "INVALID_RADIUS"
            }
    except (ValueError, TypeError):
        return {
            "status": "error",
            "message": "半徑必須是數值",
            "error_code": "INVALID_RADIUS_TYPE"
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
    area = math.pi * radius_value * radius_value
    circumference = 2 * math.pi * radius_value
    
    return {
        "status": "success",
        "data": {
            "circle_id": "AcDbCircle:1234567890",
            "center_point": center_coords,
            "radius": radius_value,
            "layer": layer.strip(),
            "area": round(area, 2),
            "circumference": round(circumference, 2),
            "created_at": datetime.now().isoformat()
        },
        "message": f"成功繪製圓形，半徑: {radius_value}",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def draw_circle(
    center_point: list,
    radius: float,
    layer: str = "0"
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_draw_circle.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_draw_circle.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_draw_circle.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_draw_circle.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_draw_circle_integration.py
import pytest
from mcp_server_fastmcp import draw_circle

class TestDrawCircleIntegration:
    def test_draw_circle_with_real_autocad(self):
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
- [ ] 面積和周長計算正確

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