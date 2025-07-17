# PRP: scan_elements

> **功能**: 掃描 AutoCAD 圖面中的元素並提取資訊  
> **類型**: MCP 工具  
> **優先級**: 高  
> **日期**: 2025年7月16日  
> **階段**: 第二階段進階繪圖工具

## 功能概述

### 目標
掃描 AutoCAD 圖面中的所有繪圖元素，並提取其幾何屬性、位置資訊和元數據。這是第二階段進階工具的核心功能，為後續的資料庫匯出和 Odoo 整合奠定基礎。

### 使用場景
1. **圖面分析**: 工程師需要分析圖面的組成元素和數量
2. **資料擷取**: 為 BOQ 生成準備元素資訊
3. **品質檢查**: 檢查圖面元素的完整性和規範性
4. **AI 助手整合**: 透過自然語言查詢圖面元素資訊

### 成功標準
- [ ] 能夠掃描並識別所有主要圖面元素類型
- [ ] 提供完整的元素屬性資訊（幾何、位置、圖層等）
- [ ] 支援元素類型過濾功能
- [ ] 支援座標轉換和單位處理
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def scan_elements(
    element_type: str = "all",
    include_geometry: bool = True,
    include_properties: bool = True,
    layer_filter: str = None,
    bounds: list = None
) -> Dict[str, Any]:
    """掃描 AutoCAD 圖面中的元素並提取資訊"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| element_type | str | 否 | "all" | 元素類型過濾 ("all", "line", "circle", "arc", "text", "dimension", "block") |
| include_geometry | bool | 否 | True | 是否包含幾何資訊 |
| include_properties | bool | 否 | True | 是否包含屬性資訊 |
| layer_filter | str | 否 | None | 圖層過濾器 |
| bounds | list | 否 | None | 掃描範圍 [x1, y1, x2, y2] |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "elements": [
            {
                "id": "AcDbLine:1234567890",
                "type": "line",
                "layer": "0",
                "color": 7,
                "geometry": {
                    "start_point": [0, 0, 0],
                    "end_point": [100, 100, 0],
                    "length": 141.42,
                    "angle": 45.0
                },
                "properties": {
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "visible": True,
                    "locked": False
                },
                "bounds": {
                    "min": [0, 0, 0],
                    "max": [100, 100, 0]
                }
            },
            {
                "id": "AcDbCircle:1234567891",
                "type": "circle",
                "layer": "0",
                "color": 7,
                "geometry": {
                    "center": [50, 50, 0],
                    "radius": 25.0,
                    "area": 1963.50,
                    "circumference": 157.08
                },
                "properties": {
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "visible": True,
                    "locked": False
                },
                "bounds": {
                    "min": [25, 25, 0],
                    "max": [75, 75, 0]
                }
            }
        ],
        "summary": {
            "total_count": 2,
            "element_counts": {
                "line": 1,
                "circle": 1
            },
            "layers": ["0"],
            "bounds": {
                "min": [0, 0, 0],
                "max": [100, 100, 0]
            }
        },
        "scan_settings": {
            "element_type": "all",
            "include_geometry": True,
            "include_properties": True,
            "layer_filter": None,
            "bounds": None
        },
        "scanned_at": "2025-07-16T10:30:00"
    },
    "message": "成功掃描 2 個圖面元素",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "無效的元素類型",
    "error_code": "INVALID_ELEMENT_TYPE",
    "suggestion": "元素類型必須是 'all', 'line', 'circle', 'arc', 'text', 'dimension', 'block' 之一"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [ ] 需要幾何計算庫（math, numpy 可選）
- [ ] 需要 win32com.client 和 pythoncom
- [ ] 需要新的 AutoCAD 元素掃描工具類別

### 核心邏輯
1. **參數驗證**: 檢查元素類型、邊界範圍格式
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **元素掃描**: 遍歷圖面空間中的所有或指定類型元素
4. **幾何資訊提取**: 計算元素的幾何屬性（長度、面積、角度等）
5. **屬性收集**: 收集元素的視覺屬性和元數據
6. **過濾處理**: 根據元素類型、圖層、邊界進行過濾
7. **資料整理**: 組織掃描結果為標準格式
8. **結果回傳**: 返回完整的掃描結果和統計資訊

### 錯誤處理
- **參數錯誤**: 元素類型無效、邊界格式錯誤
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **掃描錯誤**: 元素存取失敗、屬性讀取失敗
- **幾何計算錯誤**: 座標轉換失敗、數學計算異常

## 範例代碼

### 基本實現
```python
@mcp.tool()
def scan_elements(
    element_type: str = "all",
    include_geometry: bool = True,
    include_properties: bool = True,
    layer_filter: str = None,
    bounds: list = None
) -> Dict[str, Any]:
    """掃描 AutoCAD 圖面中的元素並提取資訊"""
    logger.info(f"scan_elements called with params: {locals()}")
    
    try:
        # 參數驗證
        valid_types = ["all", "line", "circle", "arc", "text", "dimension", "block"]
        if element_type not in valid_types:
            return {
                "status": "error",
                "message": "無效的元素類型",
                "error_code": "INVALID_ELEMENT_TYPE",
                "suggestion": f"元素類型必須是 {valid_types} 之一"
            }
        
        # 邊界驗證
        if bounds is not None:
            if not isinstance(bounds, list) or len(bounds) != 4:
                return {
                    "status": "error",
                    "message": "邊界格式無效",
                    "error_code": "INVALID_BOUNDS",
                    "suggestion": "邊界必須是 [x1, y1, x2, y2] 格式的列表"
                }
            
            try:
                bounds = [float(b) for b in bounds]
            except (ValueError, TypeError):
                return {
                    "status": "error",
                    "message": "邊界座標必須是數字",
                    "error_code": "INVALID_BOUNDS_VALUES"
                }
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 掃描元素
        result = autocad_util.scan_elements(
            element_type=element_type,
            include_geometry=include_geometry,
            include_properties=include_properties,
            layer_filter=layer_filter,
            bounds=bounds
        )
        
        return {
            "status": "success",
            "data": {
                "elements": result.get("elements", []),
                "summary": result.get("summary", {}),
                "scan_settings": {
                    "element_type": element_type,
                    "include_geometry": include_geometry,
                    "include_properties": include_properties,
                    "layer_filter": layer_filter,
                    "bounds": bounds
                },
                "scanned_at": datetime.now().isoformat()
            },
            "message": f"成功掃描 {result.get('summary', {}).get('total_count', 0)} 個圖面元素",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in scan_elements: {e}")
        return {
            "status": "error",
            "message": f"掃描元素失敗: {str(e)}",
            "error_code": "ELEMENT_SCAN_ERROR"
        }
```

### 使用範例
```python
# 基本使用 - 掃描所有元素
result = scan_elements()

# 只掃描線段
result = scan_elements(element_type="line")

# 掃描指定圖層的圓形
result = scan_elements(element_type="circle", layer_filter="CIRCLES")

# 掃描指定區域內的所有元素
result = scan_elements(bounds=[0, 0, 100, 100])

# 簡化掃描 - 不包含幾何和屬性詳細資訊
result = scan_elements(include_geometry=False, include_properties=False)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_scan_elements.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
import math

import mcp_server_fastmcp

class TestScanElements:
    def test_scan_elements_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbLine:1234567890",
                    "type": "line",
                    "layer": "0",
                    "color": 7,
                    "geometry": {
                        "start_point": [0, 0, 0],
                        "end_point": [100, 100, 0],
                        "length": 141.42,
                        "angle": 45.0
                    },
                    "properties": {
                        "linetype": "Continuous",
                        "lineweight": "Default",
                        "visible": True,
                        "locked": False
                    },
                    "bounds": {
                        "min": [0, 0, 0],
                        "max": [100, 100, 0]
                    }
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"line": 1},
                "layers": ["0"],
                "bounds": {
                    "min": [0, 0, 0],
                    "max": [100, 100, 0]
                }
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        assert len(result["data"]["elements"]) == 1
        assert result["data"]["summary"]["total_count"] == 1
        assert result["data"]["scan_settings"]["element_type"] == "all"
        assert result["data"]["scan_settings"]["include_geometry"] == True
        assert result["data"]["scan_settings"]["include_properties"] == True
        assert "scanned_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_scan_elements_parameter_validation_invalid_type(self):
        """參數驗證測試 - 無效元素類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.scan_elements(element_type="invalid")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ELEMENT_TYPE"
        
    def test_scan_elements_parameter_validation_invalid_bounds(self):
        """參數驗證測試 - 無效邊界"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.scan_elements(bounds=[1, 2, 3])  # 只有3個值
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_BOUNDS"
        
    def test_scan_elements_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.scan_elements()
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_scan_elements_with_element_type_filter(self):
        """元素類型過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbCircle:1234567891",
                    "type": "circle",
                    "layer": "0",
                    "color": 7,
                    "geometry": {
                        "center": [50, 50, 0],
                        "radius": 25.0,
                        "area": 1963.50,
                        "circumference": 157.08
                    },
                    "properties": {
                        "linetype": "Continuous",
                        "lineweight": "Default",
                        "visible": True,
                        "locked": False
                    },
                    "bounds": {
                        "min": [25, 25, 0],
                        "max": [75, 75, 0]
                    }
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"circle": 1},
                "layers": ["0"],
                "bounds": {
                    "min": [25, 25, 0],
                    "max": [75, 75, 0]
                }
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(element_type="circle")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["element_type"] == "circle"
        assert result["data"]["elements"][0]["type"] == "circle"
        
    def test_scan_elements_with_layer_filter(self):
        """圖層過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {
                "total_count": 0,
                "element_counts": {},
                "layers": [],
                "bounds": None
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(layer_filter="WALLS")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["layer_filter"] == "WALLS"
        
    def test_scan_elements_with_bounds_filter(self):
        """邊界過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {
                "total_count": 0,
                "element_counts": {},
                "layers": [],
                "bounds": None
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(bounds=[0, 0, 100, 100])
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["bounds"] == [0, 0, 100, 100]
        
    def test_scan_elements_without_geometry(self):
        """不包含幾何資訊測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbLine:1234567890",
                    "type": "line",
                    "layer": "0",
                    "color": 7
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"line": 1},
                "layers": ["0"],
                "bounds": None
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(include_geometry=False)
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["include_geometry"] == False
        
    def test_scan_elements_without_properties(self):
        """不包含屬性資訊測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbLine:1234567890",
                    "type": "line",
                    "layer": "0",
                    "color": 7
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"line": 1},
                "layers": ["0"],
                "bounds": None
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(include_properties=False)
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["include_properties"] == False
        
    def test_scan_elements_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {
                "total_count": 0,
                "element_counts": {},
                "layers": [],
                "bounds": None
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(
            element_type="line",
            include_geometry=False,
            include_properties=False,
            layer_filter="WALLS",
            bounds=[0, 0, 100, 100]
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.scan_elements.assert_called_once_with(
            element_type="line",
            include_geometry=False,
            include_properties=False,
            layer_filter="WALLS",
            bounds=[0, 0, 100, 100]
        )
        
    def test_scan_elements_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.side_effect = Exception("Element scan error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.scan_elements()
        assert result["status"] == "error"
        assert result["error_code"] == "ELEMENT_SCAN_ERROR"
        
    def test_scan_elements_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'scan_elements')
        assert callable(mcp_server_fastmcp.scan_elements)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def scan_elements(
    element_type: str = "all",
    include_geometry: bool = True,
    include_properties: bool = True,
    layer_filter: str = None,
    bounds: list = None
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    valid_types = ["all", "line", "circle", "arc", "text", "dimension", "block"]
    if element_type not in valid_types:
        return {
            "status": "error",
            "message": "無效的元素類型",
            "error_code": "INVALID_ELEMENT_TYPE"
        }
    
    if bounds is not None:
        if not isinstance(bounds, list) or len(bounds) != 4:
            return {
                "status": "error",
                "message": "邊界格式無效",
                "error_code": "INVALID_BOUNDS"
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
            "elements": [],
            "summary": {
                "total_count": 0,
                "element_counts": {},
                "layers": [],
                "bounds": None
            },
            "scan_settings": {
                "element_type": element_type,
                "include_geometry": include_geometry,
                "include_properties": include_properties,
                "layer_filter": layer_filter,
                "bounds": bounds
            },
            "scanned_at": datetime.now().isoformat()
        },
        "message": f"成功掃描 0 個圖面元素",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def scan_elements(
    element_type: str = "all",
    include_geometry: bool = True,
    include_properties: bool = True,
    layer_filter: str = None,
    bounds: list = None
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_scan_elements.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_scan_elements.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_scan_elements.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_scan_elements.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_scan_elements_integration.py
import pytest
from mcp_server_fastmcp import scan_elements

class TestScanElementsIntegration:
    def test_scan_elements_with_real_autocad(self):
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

## 實施相關

### 新增 AutoCAD 工具類別
需要在 `utility/util_autocad.py` 中新增 `scan_elements` 方法：

```python
def scan_elements(self, element_type="all", include_geometry=True, 
                 include_properties=True, layer_filter=None, bounds=None):
    """掃描圖面元素"""
    try:
        # 確保 AutoCAD 連接
        if not self.app or not self.doc:
            raise Exception("AutoCAD 連接未建立")
        
        # 獲取模型空間
        model_space = self.doc.ModelSpace
        elements = []
        element_counts = {}
        layers = set()
        all_bounds = []
        
        # 遍歷所有元素
        for entity in model_space:
            entity_type = self._get_entity_type(entity)
            
            # 元素類型過濾
            if element_type != "all" and entity_type != element_type:
                continue
                
            # 圖層過濾
            if layer_filter and entity.Layer != layer_filter:
                continue
                
            # 邊界過濾
            if bounds:
                entity_bounds = self._get_entity_bounds(entity)
                if not self._is_within_bounds(entity_bounds, bounds):
                    continue
            
            # 提取元素資訊
            element_info = {
                "id": f"AcDb{entity_type.title()}:{entity.Handle}",
                "type": entity_type,
                "layer": entity.Layer,
                "color": entity.Color
            }
            
            # 添加幾何資訊
            if include_geometry:
                element_info["geometry"] = self._extract_geometry(entity)
                
            # 添加屬性資訊
            if include_properties:
                element_info["properties"] = self._extract_properties(entity)
                
            # 添加邊界資訊
            element_info["bounds"] = self._get_entity_bounds(entity)
            
            elements.append(element_info)
            
            # 更新統計
            element_counts[entity_type] = element_counts.get(entity_type, 0) + 1
            layers.add(entity.Layer)
            all_bounds.append(element_info["bounds"])
        
        # 計算整體邊界
        overall_bounds = self._calculate_overall_bounds(all_bounds) if all_bounds else None
        
        return {
            "elements": elements,
            "summary": {
                "total_count": len(elements),
                "element_counts": element_counts,
                "layers": list(layers),
                "bounds": overall_bounds
            }
        }
        
    except Exception as e:
        raise Exception(f"掃描元素失敗: {str(e)}")
```

### 支援方法
```python
def _get_entity_type(self, entity):
    """獲取實體類型"""
    type_name = entity.ObjectName
    if type_name == "AcDbLine":
        return "line"
    elif type_name == "AcDbCircle":
        return "circle"
    elif type_name == "AcDbArc":
        return "arc"
    elif type_name == "AcDbText":
        return "text"
    elif type_name.startswith("AcDbDimension"):
        return "dimension"
    elif type_name == "AcDbBlockReference":
        return "block"
    else:
        return "unknown"

def _extract_geometry(self, entity):
    """提取幾何資訊"""
    entity_type = self._get_entity_type(entity)
    
    if entity_type == "line":
        return {
            "start_point": list(entity.StartPoint),
            "end_point": list(entity.EndPoint),
            "length": entity.Length,
            "angle": math.degrees(entity.Angle)
        }
    elif entity_type == "circle":
        return {
            "center": list(entity.Center),
            "radius": entity.Radius,
            "area": math.pi * entity.Radius * entity.Radius,
            "circumference": 2 * math.pi * entity.Radius
        }
    # ... 其他幾何類型
    
def _extract_properties(self, entity):
    """提取屬性資訊"""
    return {
        "linetype": entity.Linetype,
        "lineweight": entity.Lineweight,
        "visible": entity.Visible,
        "locked": entity.Layer in self._get_locked_layers()
    }

def _get_entity_bounds(self, entity):
    """獲取實體邊界"""
    try:
        min_point = entity.GetBoundingBox()[0]
        max_point = entity.GetBoundingBox()[1]
        return {
            "min": list(min_point),
            "max": list(max_point)
        }
    except:
        return None

def _is_within_bounds(self, entity_bounds, bounds):
    """檢查實體是否在指定邊界內"""
    if not entity_bounds:
        return True
    
    return (entity_bounds["min"][0] >= bounds[0] and
            entity_bounds["min"][1] >= bounds[1] and
            entity_bounds["max"][0] <= bounds[2] and
            entity_bounds["max"][1] <= bounds[3])

def _calculate_overall_bounds(self, all_bounds):
    """計算整體邊界"""
    if not all_bounds:
        return None
    
    min_x = min(b["min"][0] for b in all_bounds if b)
    min_y = min(b["min"][1] for b in all_bounds if b)
    min_z = min(b["min"][2] for b in all_bounds if b)
    max_x = max(b["max"][0] for b in all_bounds if b)
    max_y = max(b["max"][1] for b in all_bounds if b)
    max_z = max(b["max"][2] for b in all_bounds if b)
    
    return {
        "min": [min_x, min_y, min_z],
        "max": [max_x, max_y, max_z]
    }
```

## 驗證標準

### 功能驗證
- [ ] 基本功能正常運作
- [ ] 參數驗證正確
- [ ] 錯誤處理完整
- [ ] 回傳值格式正確
- [ ] 支援所有指定的參數組合
- [ ] 與 AutoCAD 正常互動
- [ ] 元素類型識別正確
- [ ] 幾何計算準確
- [ ] 屬性提取完整
- [ ] 過濾功能正確

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