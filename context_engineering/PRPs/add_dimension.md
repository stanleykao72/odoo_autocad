# PRP: add_dimension

> **功能**: 在 AutoCAD 中添加尺寸標註  
> **類型**: MCP 工具  
> **優先級**: 中  
> **日期**: 2025年7月16日  
> **階段**: 第二階段進階繪圖工具

## 功能概述

### 目標
在 AutoCAD 圖面中添加尺寸標註，支援線性尺寸、角度尺寸、徑向尺寸等多種尺寸類型。這是第二階段進階繪圖工具的重要功能，為圖面標註和測量提供支援。

### 使用場景
1. **工程製圖**: 為圖面元素添加準確的尺寸標註
2. **測量標記**: 標記關鍵尺寸和距離
3. **設計驗證**: 檢查設計是否符合規範
4. **施工圖面**: 提供施工所需的尺寸資訊
5. **AI 助手整合**: 透過自然語言指令創建尺寸標註

### 成功標準
- [ ] 能夠創建線性尺寸標註
- [ ] 支援角度尺寸標註
- [ ] 支援徑向尺寸標註
- [ ] 支援自訂尺寸文字和格式
- [ ] 支援不同的尺寸樣式和圖層
- [ ] 提供尺寸資訊和位置數據
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def add_dimension(
    dimension_type: str,
    definition_points: list,
    text_position: list = None,
    text_override: str = None,
    dim_style: str = "Standard",
    layer: str = "0",
    angle: float = 0.0
) -> Dict[str, Any]:
    """在 AutoCAD 中添加尺寸標註"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| dimension_type | str | 是 | - | 尺寸類型 ("linear", "angular", "radial", "diameter") |
| definition_points | list | 是 | - | 定義點列表 |
| text_position | list | 否 | None | 尺寸文字位置 [x, y, z] |
| text_override | str | 否 | None | 自訂尺寸文字 |
| dim_style | str | 否 | "Standard" | 尺寸樣式 |
| layer | str | 否 | "0" | 圖層名稱 |
| angle | float | 否 | 0.0 | 尺寸角度（度） |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "dimension_id": "AcDbDimension:1234567890",
        "dimension_type": "linear",
        "definition_points": [[0, 0, 0], [100, 0, 0]],
        "text_position": [50, 10, 0],
        "measured_value": 100.0,
        "display_text": "100.00",
        "text_override": null,
        "dim_style": "Standard",
        "layer": "0",
        "angle": 0.0,
        "properties": {
            "color": 7,
            "linetype": "Continuous",
            "lineweight": "Default",
            "visible": True,
            "locked": False
        },
        "dimension_info": {
            "units": "mm",
            "precision": 2,
            "scale": 1.0,
            "arrow_size": 2.5,
            "text_height": 2.5
        },
        "bounds": {
            "min": [0, 0, 0],
            "max": [100, 15, 0],
            "width": 100.0,
            "height": 15.0
        },
        "created_at": "2025-07-16T10:30:00"
    },
    "message": "成功添加線性尺寸標註: 100.00",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "無效的尺寸類型",
    "error_code": "INVALID_DIMENSION_TYPE",
    "suggestion": "尺寸類型必須是 'linear', 'angular', 'radial', 'diameter' 之一"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [x] 需要 win32com.client 和 pythoncom
- [ ] 需要尺寸樣式管理功能
- [ ] 需要測量計算功能

### 核心邏輯
1. **參數驗證**: 檢查尺寸類型、定義點、文字位置
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **尺寸類型處理**: 根據類型創建相應的尺寸標註
4. **定義點驗證**: 確保定義點數量和格式正確
5. **尺寸創建**: 使用 AutoCAD COM API 創建尺寸物件
6. **屬性設定**: 設定尺寸的樣式、圖層、角度等屬性
7. **測量計算**: 計算實際測量值
8. **結果回傳**: 返回創建的尺寸物件資訊

### 錯誤處理
- **參數錯誤**: 尺寸類型無效、定義點不足、文字位置格式無效
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **創建錯誤**: 尺寸創建失敗、屬性設定失敗
- **樣式錯誤**: 尺寸樣式不存在、無法套用樣式

## 範例代碼

### 基本實現
```python
@mcp.tool()
def add_dimension(
    dimension_type: str,
    definition_points: list,
    text_position: list = None,
    text_override: str = None,
    dim_style: str = "Standard",
    layer: str = "0",
    angle: float = 0.0
) -> Dict[str, Any]:
    """在 AutoCAD 中添加尺寸標註"""
    logger.info(f"add_dimension called with params: {locals()}")
    
    try:
        # 參數驗證
        valid_types = ["linear", "angular", "radial", "diameter"]
        if dimension_type not in valid_types:
            return {
                "status": "error",
                "message": f"無效的尺寸類型: {dimension_type}",
                "error_code": "INVALID_DIMENSION_TYPE",
                "suggestion": f"尺寸類型必須是 {valid_types} 之一"
            }
        
        if not isinstance(definition_points, list) or len(definition_points) < 2:
            return {
                "status": "error",
                "message": "定義點必須至少包含兩個點",
                "error_code": "INVALID_DEFINITION_POINTS",
                "suggestion": "請提供至少兩個定義點的座標"
            }
        
        # 驗證定義點格式
        processed_points = []
        for i, point in enumerate(definition_points):
            if not isinstance(point, list) or len(point) != 3:
                return {
                    "status": "error",
                    "message": f"定義點 {i+1} 格式無效",
                    "error_code": "INVALID_POINT_FORMAT",
                    "suggestion": "每個定義點必須是 [x, y, z] 格式"
                }
            try:
                processed_points.append([float(p) for p in point])
            except (ValueError, TypeError):
                return {
                    "status": "error",
                    "message": f"定義點 {i+1} 座標必須是數字",
                    "error_code": "INVALID_POINT_VALUES"
                }
        
        # 處理文字位置
        if text_position is not None:
            if not isinstance(text_position, list) or len(text_position) != 3:
                return {
                    "status": "error",
                    "message": "文字位置格式無效",
                    "error_code": "INVALID_TEXT_POSITION",
                    "suggestion": "文字位置必須是 [x, y, z] 格式"
                }
            try:
                text_position = [float(p) for p in text_position]
            except (ValueError, TypeError):
                return {
                    "status": "error",
                    "message": "文字位置座標必須是數字",
                    "error_code": "INVALID_TEXT_POSITION_VALUES"
                }
        
        # 檢查角度
        try:
            angle = float(angle)
            angle_rad = math.radians(angle)
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "角度必須是數字",
                "error_code": "INVALID_ANGLE_TYPE"
            }
        
        # 檢查圖層
        if not layer or not layer.strip():
            return {
                "status": "error",
                "message": "圖層名稱不能為空",
                "error_code": "INVALID_LAYER"
            }
        
        layer = layer.strip()
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 創建尺寸
        result = autocad_util.add_dimension(
            dimension_type=dimension_type,
            definition_points=processed_points,
            text_position=text_position,
            text_override=text_override,
            dim_style=dim_style,
            layer=layer,
            angle=angle_rad
        )
        
        return {
            "status": "success",
            "data": {
                "dimension_id": result.get("dimension_id"),
                "dimension_type": dimension_type,
                "definition_points": processed_points,
                "text_position": text_position,
                "measured_value": result.get("measured_value", 0.0),
                "display_text": result.get("display_text", ""),
                "text_override": text_override,
                "dim_style": dim_style,
                "layer": layer,
                "angle": angle,
                "properties": result.get("properties", {}),
                "dimension_info": result.get("dimension_info", {}),
                "bounds": result.get("bounds", {}),
                "created_at": datetime.now().isoformat()
            },
            "message": f"成功添加{dimension_type}尺寸標註: {result.get('display_text', '')}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in add_dimension: {e}")
        return {
            "status": "error",
            "message": f"添加尺寸標註失敗: {str(e)}",
            "error_code": "DIMENSION_CREATION_ERROR"
        }
```

### AutoCAD 工具類別方法
```python
def add_dimension(self, dimension_type, definition_points, text_position, 
                 text_override, dim_style, layer, angle):
    """在 AutoCAD 中添加尺寸標註"""
    try:
        # 確保 AutoCAD 連接
        if not self.acad or not self.doc:
            raise Exception("AutoCAD 連接未建立")
        
        # 確保圖層存在
        self._ensure_layer_exists(layer)
        
        # 獲取模型空間
        model_space = self.doc.ModelSpace
        
        # 根據尺寸類型創建相應的尺寸
        if dimension_type == "linear":
            # 線性尺寸需要兩個定義點和一個文字位置
            if len(definition_points) < 2:
                raise Exception("線性尺寸需要至少兩個定義點")
            
            point1 = definition_points[0]
            point2 = definition_points[1]
            
            # 如果沒有指定文字位置，計算預設位置
            if text_position is None:
                mid_x = (point1[0] + point2[0]) / 2
                mid_y = (point1[1] + point2[1]) / 2 + 10  # 向上偏移10單位
                text_position = [mid_x, mid_y, 0]
            
            dim_obj = model_space.AddDimAligned(point1, point2, text_position)
            measured_value = abs(point2[0] - point1[0]) if abs(point2[0] - point1[0]) > abs(point2[1] - point1[1]) else abs(point2[1] - point1[1])
            
        elif dimension_type == "angular":
            # 角度尺寸需要三個定義點
            if len(definition_points) < 3:
                raise Exception("角度尺寸需要至少三個定義點")
            
            center = definition_points[0]
            point1 = definition_points[1]
            point2 = definition_points[2]
            
            if text_position is None:
                # 計算角度中點作為文字位置
                text_position = [center[0] + 20, center[1] + 20, 0]
            
            dim_obj = model_space.AddDimAngular(center, point1, point2, text_position)
            # 計算角度
            import math
            angle1 = math.atan2(point1[1] - center[1], point1[0] - center[0])
            angle2 = math.atan2(point2[1] - center[1], point2[0] - center[0])
            measured_value = abs(math.degrees(angle2 - angle1))
            
        elif dimension_type == "radial":
            # 徑向尺寸需要圓心和圓上一點
            if len(definition_points) < 2:
                raise Exception("徑向尺寸需要至少兩個定義點")
            
            center = definition_points[0]
            point_on_circle = definition_points[1]
            
            if text_position is None:
                # 計算徑向文字位置
                text_position = [(center[0] + point_on_circle[0]) / 2, 
                               (center[1] + point_on_circle[1]) / 2, 0]
            
            dim_obj = model_space.AddDimRadial(center, point_on_circle, text_position)
            # 計算半徑
            measured_value = ((point_on_circle[0] - center[0])**2 + 
                            (point_on_circle[1] - center[1])**2)**0.5
            
        elif dimension_type == "diameter":
            # 直徑尺寸需要兩個對角點
            if len(definition_points) < 2:
                raise Exception("直徑尺寸需要至少兩個定義點")
            
            point1 = definition_points[0]
            point2 = definition_points[1]
            
            if text_position is None:
                text_position = [(point1[0] + point2[0]) / 2, 
                               (point1[1] + point2[1]) / 2, 0]
            
            dim_obj = model_space.AddDimDiametric(point1, point2, text_position)
            measured_value = ((point2[0] - point1[0])**2 + 
                            (point2[1] - point1[1])**2)**0.5
            
        # 設定屬性
        dim_obj.Layer = layer
        
        # 設定尺寸樣式
        try:
            dim_obj.StyleName = dim_style
        except Exception:
            # 如果樣式不存在，使用預設樣式
            dim_obj.StyleName = "Standard"
        
        # 設定自訂文字
        if text_override:
            dim_obj.TextOverride = text_override
        
        # 獲取尺寸屬性
        properties = {
            "color": dim_obj.Color,
            "linetype": dim_obj.Linetype,
            "lineweight": dim_obj.Lineweight,
            "visible": dim_obj.Visible,
            "locked": False
        }
        
        # 獲取尺寸資訊
        dimension_info = {
            "units": "mm",  # 預設單位
            "precision": 2,
            "scale": 1.0,
            "arrow_size": 2.5,
            "text_height": 2.5
        }
        
        # 計算邊界
        try:
            bounds_min = dim_obj.GetBoundingBox()[0]
            bounds_max = dim_obj.GetBoundingBox()[1]
            bounds = {
                "min": list(bounds_min),
                "max": list(bounds_max),
                "width": bounds_max[0] - bounds_min[0],
                "height": bounds_max[1] - bounds_min[1]
            }
        except Exception:
            # 如果無法獲取邊界，使用估算值
            bounds = {
                "min": text_position,
                "max": [text_position[0] + 50, text_position[1] + 15, text_position[2]],
                "width": 50.0,
                "height": 15.0
            }
        
        # 獲取顯示文字
        try:
            display_text = dim_obj.TextString
        except Exception:
            display_text = f"{measured_value:.2f}"
        
        self.log.safe_log_insert(f"成功添加{dimension_type}尺寸標註: {display_text}\n")
        
        return {
            "dimension_id": f"AcDbDimension:{dim_obj.Handle}",
            "measured_value": measured_value,
            "display_text": display_text,
            "properties": properties,
            "dimension_info": dimension_info,
            "bounds": bounds
        }
        
    except Exception as e:
        self.log.safe_log_insert(f"添加尺寸標註時發生錯誤: {str(e)}\n")
        raise e
```

### 使用範例
```python
# 線性尺寸
result = add_dimension(
    dimension_type="linear",
    definition_points=[[0, 0, 0], [100, 0, 0]]
)

# 角度尺寸
result = add_dimension(
    dimension_type="angular",
    definition_points=[[0, 0, 0], [50, 0, 0], [50, 50, 0]]
)

# 徑向尺寸
result = add_dimension(
    dimension_type="radial",
    definition_points=[[0, 0, 0], [25, 0, 0]],
    text_position=[15, 15, 0]
)

# 直徑尺寸
result = add_dimension(
    dimension_type="diameter",
    definition_points=[[-25, 0, 0], [25, 0, 0]]
)

# 自訂文字的尺寸
result = add_dimension(
    dimension_type="linear",
    definition_points=[[0, 0, 0], [100, 0, 0]],
    text_override="100mm",
    dim_style="ISO-25",
    layer="DIMENSIONS"
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_add_dimension.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
import math

import mcp_server_fastmcp

class TestAddDimension:
    def test_add_dimension_linear_basic_functionality(self):
        """線性尺寸基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {
                "color": 7,
                "linetype": "Continuous",
                "lineweight": "Default",
                "visible": True,
                "locked": False
            },
            "dimension_info": {
                "units": "mm",
                "precision": 2,
                "scale": 1.0,
                "arrow_size": 2.5,
                "text_height": 2.5
            },
            "bounds": {
                "min": [0, 0, 0],
                "max": [100, 15, 0],
                "width": 100.0,
                "height": 15.0
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["dimension_type"] == "linear"
        assert result["data"]["definition_points"] == [[0.0, 0.0, 0.0], [100.0, 0.0, 0.0]]
        assert result["data"]["measured_value"] == 100.0
        assert result["data"]["display_text"] == "100.00"
        assert "properties" in result["data"]
        assert "dimension_info" in result["data"]
        assert "bounds" in result["data"]
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_add_dimension_parameter_validation_invalid_type(self):
        """參數驗證測試 - 無效尺寸類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="invalid_type",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DIMENSION_TYPE"
        
    def test_add_dimension_parameter_validation_insufficient_points(self):
        """參數驗證測試 - 定義點不足"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0]]  # 只有一個點
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DEFINITION_POINTS"
        
    def test_add_dimension_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_add_dimension_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'add_dimension')
        assert callable(mcp_server_fastmcp.add_dimension)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def add_dimension(
    dimension_type: str,
    definition_points: list,
    text_position: list = None,
    text_override: str = None,
    dim_style: str = "Standard",
    layer: str = "0",
    angle: float = 0.0
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    valid_types = ["linear", "angular", "radial", "diameter"]
    if dimension_type not in valid_types:
        return {
            "status": "error",
            "message": f"無效的尺寸類型: {dimension_type}",
            "error_code": "INVALID_DIMENSION_TYPE"
        }
    
    if not isinstance(definition_points, list) or len(definition_points) < 2:
        return {
            "status": "error",
            "message": "定義點必須至少包含兩個點",
            "error_code": "INVALID_DEFINITION_POINTS"
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
            "dimension_id": "AcDbDimension:1234567890",
            "dimension_type": dimension_type,
            "definition_points": [[float(p) for p in point] for point in definition_points],
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {},
            "created_at": datetime.now().isoformat()
        },
        "message": f"成功添加{dimension_type}尺寸標註",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def add_dimension(
    dimension_type: str,
    definition_points: list,
    text_position: list = None,
    text_override: str = None,
    dim_style: str = "Standard",
    layer: str = "0",
    angle: float = 0.0
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_add_dimension.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_add_dimension.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_add_dimension.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_add_dimension.py --cov=mcp_server_fastmcp --cov-report=term-missing
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
- [ ] 支援所有指定的尺寸類型
- [ ] 與 AutoCAD 正常互動
- [ ] 尺寸標註位置正確
- [ ] 尺寸屬性設定正確
- [ ] 測量值計算準確
- [ ] 邊界計算準確

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
- [ ] 與其他繪圖工具相容
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
- [尺寸標註模式](../examples/dimension_patterns.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_autocad.py` - AutoCAD 工具類別
- `tests/unit/test_draw_*.py` - 相關測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。