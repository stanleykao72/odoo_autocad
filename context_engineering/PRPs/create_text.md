# PRP: create_text

> **功能**: 在 AutoCAD 中創建文字註解  
> **類型**: MCP 工具  
> **優先級**: 中  
> **日期**: 2025年7月16日  
> **階段**: 第二階段進階繪圖工具

## 功能概述

### 目標
在 AutoCAD 圖面中創建文字註解，支援自訂位置、文字內容、高度、旋轉角度、圖層等屬性。這是第二階段進階繪圖工具的重要功能，為圖面標註和說明提供支援。

### 使用場景
1. **圖面標註**: 工程師需要在圖面上添加說明文字
2. **尺寸標記**: 為圖面元素添加尺寸或編號標記
3. **圖例說明**: 創建圖例和說明文字
4. **工程註解**: 添加工程相關的技術註解
5. **AI 助手整合**: 透過自然語言指令創建文字註解

### 成功標準
- [ ] 能夠在指定位置創建文字註解
- [ ] 支援文字內容、高度、旋轉角度的自訂
- [ ] 支援圖層、樣式、對齊方式的設定
- [ ] 提供文字邊界和屬性資訊
- [ ] 支援中文和英文文字
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def create_text(
    position: list,
    text_content: str,
    height: float = 2.5,
    rotation: float = 0.0,
    layer: str = "0",
    style: str = "Standard",
    alignment: str = "left"
) -> Dict[str, Any]:
    """在 AutoCAD 中創建文字註解"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| position | list | 是 | - | 文字位置 [x, y, z] |
| text_content | str | 是 | - | 文字內容 |
| height | float | 否 | 2.5 | 文字高度 |
| rotation | float | 否 | 0.0 | 旋轉角度（度） |
| layer | str | 否 | "0" | 圖層名稱 |
| style | str | 否 | "Standard" | 文字樣式 |
| alignment | str | 否 | "left" | 對齊方式 ("left", "center", "right") |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "text_id": "AcDbText:1234567890",
        "position": [100, 50, 0],
        "text_content": "示例文字",
        "height": 2.5,
        "rotation": 0.0,
        "layer": "0",
        "style": "Standard",
        "alignment": "left",
        "properties": {
            "color": 7,
            "linetype": "Continuous",
            "lineweight": "Default",
            "visible": True,
            "locked": False
        },
        "bounds": {
            "min": [100, 50, 0],
            "max": [120, 52.5, 0],
            "width": 20.0,
            "height": 2.5
        },
        "text_info": {
            "character_count": 4,
            "line_count": 1,
            "font_name": "Arial",
            "is_bold": False,
            "is_italic": False
        },
        "created_at": "2025-07-16T10:30:00"
    },
    "message": "成功創建文字註解: 示例文字",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "文字位置格式無效",
    "error_code": "INVALID_POSITION",
    "suggestion": "位置必須是 [x, y, z] 格式的數字列表"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [x] 需要 win32com.client 和 pythoncom
- [ ] 需要文字樣式管理功能
- [ ] 需要字型支援（特別是中文字型）

### 核心邏輯
1. **參數驗證**: 檢查位置座標、文字內容、高度、旋轉角度
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **圖層處理**: 確保目標圖層存在，必要時創建
4. **樣式檢查**: 驗證文字樣式是否存在
5. **文字創建**: 使用 AutoCAD COM API 創建文字物件
6. **屬性設定**: 設定文字的所有屬性（位置、內容、高度、旋轉等）
7. **邊界計算**: 計算文字的邊界框
8. **結果回傳**: 返回創建的文字物件資訊

### 錯誤處理
- **參數錯誤**: 位置格式無效、文字內容為空、高度或旋轉角度無效
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **創建錯誤**: 文字創建失敗、屬性設定失敗
- **樣式錯誤**: 文字樣式不存在、字型不支援

## 範例代碼

### 基本實現
```python
@mcp.tool()
def create_text(
    position: list,
    text_content: str,
    height: float = 2.5,
    rotation: float = 0.0,
    layer: str = "0",
    style: str = "Standard",
    alignment: str = "left"
) -> Dict[str, Any]:
    """在 AutoCAD 中創建文字註解"""
    logger.info(f"create_text called with params: {locals()}")
    
    try:
        # 參數驗證
        if not isinstance(position, list) or len(position) != 3:
            return {
                "status": "error",
                "message": "文字位置格式無效",
                "error_code": "INVALID_POSITION",
                "suggestion": "位置必須是 [x, y, z] 格式的數字列表"
            }
        
        try:
            position = [float(p) for p in position]
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "位置座標必須是數字",
                "error_code": "INVALID_POSITION_VALUES"
            }
        
        if not text_content or not text_content.strip():
            return {
                "status": "error",
                "message": "文字內容不能為空",
                "error_code": "EMPTY_TEXT_CONTENT",
                "suggestion": "請提供有效的文字內容"
            }
        
        text_content = text_content.strip()
        
        try:
            height = float(height)
            if height <= 0:
                return {
                    "status": "error",
                    "message": "文字高度必須大於 0",
                    "error_code": "INVALID_HEIGHT"
                }
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "文字高度必須是數字",
                "error_code": "INVALID_HEIGHT_TYPE"
            }
        
        try:
            rotation = float(rotation)
            # 將角度轉換為弧度
            rotation_rad = math.radians(rotation)
        except (ValueError, TypeError):
            return {
                "status": "error",
                "message": "旋轉角度必須是數字",
                "error_code": "INVALID_ROTATION_TYPE"
            }
        
        if not layer or not layer.strip():
            return {
                "status": "error",
                "message": "圖層名稱不能為空",
                "error_code": "INVALID_LAYER"
            }
        
        layer = layer.strip()
        
        valid_alignments = ["left", "center", "right"]
        if alignment not in valid_alignments:
            return {
                "status": "error",
                "message": f"無效的對齊方式: {alignment}",
                "error_code": "INVALID_ALIGNMENT",
                "suggestion": f"對齊方式必須是 {valid_alignments} 之一"
            }
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 創建文字
        result = autocad_util.create_text(
            position=position,
            text_content=text_content,
            height=height,
            rotation=rotation_rad,
            layer=layer,
            style=style,
            alignment=alignment
        )
        
        return {
            "status": "success",
            "data": {
                "text_id": result.get("text_id"),
                "position": position,
                "text_content": text_content,
                "height": height,
                "rotation": rotation,
                "layer": layer,
                "style": style,
                "alignment": alignment,
                "properties": result.get("properties", {}),
                "bounds": result.get("bounds", {}),
                "text_info": result.get("text_info", {}),
                "created_at": datetime.now().isoformat()
            },
            "message": f"成功創建文字註解: {text_content}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in create_text: {e}")
        return {
            "status": "error",
            "message": f"創建文字失敗: {str(e)}",
            "error_code": "TEXT_CREATION_ERROR"
        }
```

### AutoCAD 工具類別方法
```python
def create_text(self, position, text_content, height, rotation, layer, style, alignment):
    """在 AutoCAD 中創建文字"""
    try:
        # 確保 AutoCAD 連接
        if not self.acad or not self.doc:
            raise Exception("AutoCAD 連接未建立")
        
        # 確保圖層存在
        self._ensure_layer_exists(layer)
        
        # 獲取模型空間
        model_space = self.doc.ModelSpace
        
        # 創建文字物件
        text_obj = model_space.AddText(text_content, position, height)
        
        # 設定屬性
        text_obj.Layer = layer
        text_obj.Rotation = rotation
        
        # 設定對齊方式
        if alignment == "center":
            text_obj.Alignment = 1  # acAlignmentMiddleCenter
        elif alignment == "right":
            text_obj.Alignment = 2  # acAlignmentTopRight
        else:
            text_obj.Alignment = 0  # acAlignmentLeft
        
        # 設定文字樣式
        try:
            text_obj.StyleName = style
        except Exception:
            # 如果樣式不存在，使用預設樣式
            text_obj.StyleName = "Standard"
        
        # 獲取文字屬性
        properties = {
            "color": text_obj.Color,
            "linetype": text_obj.Linetype,
            "lineweight": text_obj.Lineweight,
            "visible": text_obj.Visible,
            "locked": False  # 文字通常不鎖定
        }
        
        # 計算文字邊界
        try:
            bounds_min = text_obj.GetBoundingBox()[0]
            bounds_max = text_obj.GetBoundingBox()[1]
            bounds = {
                "min": list(bounds_min),
                "max": list(bounds_max),
                "width": bounds_max[0] - bounds_min[0],
                "height": bounds_max[1] - bounds_min[1]
            }
        except Exception:
            # 如果無法獲取邊界，使用估算值
            estimated_width = len(text_content) * height * 0.6
            bounds = {
                "min": position,
                "max": [position[0] + estimated_width, position[1] + height, position[2]],
                "width": estimated_width,
                "height": height
            }
        
        # 獲取文字資訊
        text_info = {
            "character_count": len(text_content),
            "line_count": text_content.count('\n') + 1,
            "font_name": "Arial",  # 預設字型
            "is_bold": False,
            "is_italic": False
        }
        
        self.log.safe_log_insert(f"成功創建文字: {text_content}，位置: {position}\n")
        
        return {
            "text_id": f"AcDbText:{text_obj.Handle}",
            "properties": properties,
            "bounds": bounds,
            "text_info": text_info
        }
        
    except Exception as e:
        self.log.safe_log_insert(f"創建文字時發生錯誤: {str(e)}\n")
        raise e
```

### 使用範例
```python
# 基本使用 - 創建簡單文字
result = create_text(
    position=[100, 50, 0],
    text_content="示例文字"
)

# 創建大號文字
result = create_text(
    position=[200, 100, 0],
    text_content="標題文字",
    height=5.0
)

# 創建旋轉文字
result = create_text(
    position=[150, 75, 0],
    text_content="旋轉文字",
    rotation=45.0
)

# 創建居中對齊文字
result = create_text(
    position=[300, 150, 0],
    text_content="居中文字",
    alignment="center"
)

# 創建指定圖層的文字
result = create_text(
    position=[250, 200, 0],
    text_content="註解文字",
    layer="TEXT",
    style="Arial"
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_create_text.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
import math

import mcp_server_fastmcp

class TestCreateText:
    def test_create_text_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {
                "color": 7,
                "linetype": "Continuous",
                "lineweight": "Default",
                "visible": True,
                "locked": False
            },
            "bounds": {
                "min": [100, 50, 0],
                "max": [120, 52.5, 0],
                "width": 20.0,
                "height": 2.5
            },
            "text_info": {
                "character_count": 4,
                "line_count": 1,
                "font_name": "Arial",
                "is_bold": False,
                "is_italic": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="示例文字"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["text_id"] == "AcDbText:1234567890"
        assert result["data"]["position"] == [100, 50, 0]
        assert result["data"]["text_content"] == "示例文字"
        assert result["data"]["height"] == 2.5
        assert result["data"]["rotation"] == 0.0
        assert result["data"]["layer"] == "0"
        assert result["data"]["alignment"] == "left"
        assert "properties" in result["data"]
        assert "bounds" in result["data"]
        assert "text_info" in result["data"]
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_create_text_parameter_validation_invalid_position_length(self):
        """參數驗證測試 - 位置長度無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50],  # 只有2個座標
            text_content="測試文字"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POSITION"
        
    def test_create_text_parameter_validation_invalid_position_type(self):
        """參數驗證測試 - 位置類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position="not_a_list",
            text_content="測試文字"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POSITION"
        
    def test_create_text_parameter_validation_invalid_position_values(self):
        """參數驗證測試 - 位置值無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=["x", "y", "z"],
            text_content="測試文字"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POSITION_VALUES"
        
    def test_create_text_parameter_validation_empty_text_content(self):
        """參數驗證測試 - 文字內容為空"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content=""
        )
        assert result["status"] == "error"
        assert result["error_code"] == "EMPTY_TEXT_CONTENT"
        
    def test_create_text_parameter_validation_whitespace_text_content(self):
        """參數驗證測試 - 文字內容只有空白"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="   "
        )
        assert result["status"] == "error"
        assert result["error_code"] == "EMPTY_TEXT_CONTENT"
        
    def test_create_text_parameter_validation_invalid_height_negative(self):
        """參數驗證測試 - 高度為負數"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height=-1.0
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_HEIGHT"
        
    def test_create_text_parameter_validation_invalid_height_zero(self):
        """參數驗證測試 - 高度為零"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height=0.0
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_HEIGHT"
        
    def test_create_text_parameter_validation_invalid_height_type(self):
        """參數驗證測試 - 高度類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height="not_a_number"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_HEIGHT_TYPE"
        
    def test_create_text_parameter_validation_invalid_rotation_type(self):
        """參數驗證測試 - 旋轉角度類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            rotation="not_a_number"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ROTATION_TYPE"
        
    def test_create_text_parameter_validation_invalid_layer_empty(self):
        """參數驗證測試 - 圖層名稱為空"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            layer=""
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER"
        
    def test_create_text_parameter_validation_invalid_alignment(self):
        """參數驗證測試 - 對齊方式無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            alignment="invalid"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ALIGNMENT"
        
    def test_create_text_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_create_text_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'create_text')
        assert callable(mcp_server_fastmcp.create_text)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def create_text(
    position: list,
    text_content: str,
    height: float = 2.5,
    rotation: float = 0.0,
    layer: str = "0",
    style: str = "Standard",
    alignment: str = "left"
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    if not isinstance(position, list) or len(position) != 3:
        return {
            "status": "error",
            "message": "文字位置格式無效",
            "error_code": "INVALID_POSITION"
        }
    
    try:
        position = [float(p) for p in position]
    except (ValueError, TypeError):
        return {
            "status": "error",
            "message": "位置座標必須是數字",
            "error_code": "INVALID_POSITION_VALUES"
        }
    
    if not text_content or not text_content.strip():
        return {
            "status": "error",
            "message": "文字內容不能為空",
            "error_code": "EMPTY_TEXT_CONTENT"
        }
    
    try:
        height = float(height)
        if height <= 0:
            return {
                "status": "error",
                "message": "文字高度必須大於 0",
                "error_code": "INVALID_HEIGHT"
            }
    except (ValueError, TypeError):
        return {
            "status": "error",
            "message": "文字高度必須是數字",
            "error_code": "INVALID_HEIGHT_TYPE"
        }
    
    try:
        rotation = float(rotation)
    except (ValueError, TypeError):
        return {
            "status": "error",
            "message": "旋轉角度必須是數字",
            "error_code": "INVALID_ROTATION_TYPE"
        }
    
    if not layer or not layer.strip():
        return {
            "status": "error",
            "message": "圖層名稱不能為空",
            "error_code": "INVALID_LAYER"
        }
    
    valid_alignments = ["left", "center", "right"]
    if alignment not in valid_alignments:
        return {
            "status": "error",
            "message": f"無效的對齊方式: {alignment}",
            "error_code": "INVALID_ALIGNMENT"
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
            "text_id": "AcDbText:1234567890",
            "position": position,
            "text_content": text_content.strip(),
            "height": height,
            "rotation": rotation,
            "layer": layer.strip(),
            "style": style,
            "alignment": alignment,
            "properties": {
                "color": 7,
                "linetype": "Continuous",
                "lineweight": "Default",
                "visible": True,
                "locked": False
            },
            "bounds": {
                "min": position,
                "max": [position[0] + len(text_content) * height * 0.6, position[1] + height, position[2]],
                "width": len(text_content) * height * 0.6,
                "height": height
            },
            "text_info": {
                "character_count": len(text_content.strip()),
                "line_count": 1,
                "font_name": "Arial",
                "is_bold": False,
                "is_italic": False
            },
            "created_at": datetime.now().isoformat()
        },
        "message": f"成功創建文字註解: {text_content.strip()}",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def create_text(
    position: list,
    text_content: str,
    height: float = 2.5,
    rotation: float = 0.0,
    layer: str = "0",
    style: str = "Standard",
    alignment: str = "left"
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_create_text.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_create_text.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_create_text.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_create_text.py --cov=mcp_server_fastmcp --cov-report=term-missing
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
- [ ] 文字創建位置正確
- [ ] 文字屬性設定正確
- [ ] 支援中文文字
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
- [文字處理模式](../examples/text_handling_patterns.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_autocad.py` - AutoCAD 工具類別
- `tests/unit/test_draw_*.py` - 相關測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。