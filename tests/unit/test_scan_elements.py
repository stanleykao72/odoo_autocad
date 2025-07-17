"""
TDD 測試：scan_elements MCP 工具

遵循 Red-Green-Refactor 循環：
1. Red: 寫失敗測試 (此階段)
2. Green: 最小實現
3. Refactor: 完整實現

測試覆蓋率要求: 95%+
"""

import pytest
from unittest.mock import patch, Mock, MagicMock
from datetime import datetime
import math
import os
import sys

# 添加專案根目錄到 Python 路徑
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import mcp_server_fastmcp


class TestScanElements:
    """scan_elements 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
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
        
        # Act
        result = mcp_server_fastmcp.scan_elements(element_type="invalid")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ELEMENT_TYPE"
        assert "無效的元素類型" in result["message"]
        
    def test_scan_elements_parameter_validation_valid_types(self):
        """參數驗證測試 - 有效元素類型"""
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
        
        valid_types = ["all", "line", "circle", "arc", "text", "dimension", "block"]
        
        for element_type in valid_types:
            # Act
            result = mcp_server_fastmcp.scan_elements(element_type=element_type)
            
            # Assert
            assert result["status"] == "success"
            assert result["data"]["scan_settings"]["element_type"] == element_type
        
    def test_scan_elements_parameter_validation_invalid_bounds_length(self):
        """參數驗證測試 - 無效邊界長度"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(bounds=[1, 2, 3])  # 只有3個值
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_BOUNDS"
        assert "邊界格式無效" in result["message"]
        
    def test_scan_elements_parameter_validation_invalid_bounds_type(self):
        """參數驗證測試 - 無效邊界類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(bounds="invalid")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_BOUNDS"
        
    def test_scan_elements_parameter_validation_invalid_bounds_values(self):
        """參數驗證測試 - 無效邊界值"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(bounds=["a", "b", "c", "d"])
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_BOUNDS_VALUES"
        assert "邊界座標必須是數字" in result["message"]
        
    def test_scan_elements_parameter_validation_valid_bounds(self):
        """參數驗證測試 - 有效邊界"""
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
        assert result["data"]["scan_settings"]["bounds"] == [0.0, 0.0, 100.0, 100.0]
        
    def test_scan_elements_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
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
        assert result["data"]["scan_settings"]["bounds"] == [0.0, 0.0, 100.0, 100.0]
        
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
        
    def test_scan_elements_multiple_elements(self):
        """多元素測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbLine:1234567890",
                    "type": "line",
                    "layer": "0",
                    "color": 7
                },
                {
                    "id": "AcDbCircle:1234567891",
                    "type": "circle",
                    "layer": "CIRCLES",
                    "color": 2
                },
                {
                    "id": "AcDbText:1234567892",
                    "type": "text",
                    "layer": "TEXT",
                    "color": 3
                }
            ],
            "summary": {
                "total_count": 3,
                "element_counts": {"line": 1, "circle": 1, "text": 1},
                "layers": ["0", "CIRCLES", "TEXT"],
                "bounds": {
                    "min": [0, 0, 0],
                    "max": [200, 200, 0]
                }
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["summary"]["total_count"] == 3
        assert len(result["data"]["elements"]) == 3
        assert result["data"]["summary"]["element_counts"]["line"] == 1
        assert result["data"]["summary"]["element_counts"]["circle"] == 1
        assert result["data"]["summary"]["element_counts"]["text"] == 1
        
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
            bounds=[0.0, 0.0, 100.0, 100.0]
        )
        
    def test_scan_elements_autocad_util_called_with_defaults(self):
        """驗證使用預設值時的呼叫"""
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
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.scan_elements.assert_called_once_with(
            element_type="all",
            include_geometry=True,
            include_properties=True,
            layer_filter=None,
            bounds=None
        )
        
    def test_scan_elements_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.side_effect = Exception("Element scan error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "ELEMENT_SCAN_ERROR"
        assert "掃描元素失敗" in result["message"]
        
    def test_scan_elements_empty_result(self):
        """空結果測試"""
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
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["summary"]["total_count"] == 0
        assert len(result["data"]["elements"]) == 0
        assert result["message"] == "成功掃描 0 個圖面元素"
        
    def test_scan_elements_return_format_complete(self):
        """驗證回傳格式完整性"""
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
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "elements" in data
        assert "summary" in data
        assert "scan_settings" in data
        assert "scanned_at" in data
        
        # 檢查 scan_settings 結構
        scan_settings = data["scan_settings"]
        assert "element_type" in scan_settings
        assert "include_geometry" in scan_settings
        assert "include_properties" in scan_settings
        assert "layer_filter" in scan_settings
        assert "bounds" in scan_settings
        
        # 檢查 summary 結構
        summary = data["summary"]
        assert "total_count" in summary
        assert "element_counts" in summary
        assert "layers" in summary
        assert "bounds" in summary
        
        # 檢查時間戳格式
        assert isinstance(data["scanned_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_scan_elements_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
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
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scanned_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_scan_elements_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
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
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "scan_elements called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_scan_elements_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in scan_elements" in call_args
        
    def test_scan_elements_layer_filter_none(self):
        """圖層過濾 None 測試"""
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
        result = mcp_server_fastmcp.scan_elements(layer_filter=None)
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["layer_filter"] is None
        
    def test_scan_elements_bounds_none(self):
        """邊界過濾 None 測試"""
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
        result = mcp_server_fastmcp.scan_elements(bounds=None)
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["bounds"] is None
        
    def test_scan_elements_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [{}, {}],  # 2個元素
            "summary": {
                "total_count": 2,
                "element_counts": {"line": 1, "circle": 1},
                "layers": ["0"],
                "bounds": None
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements()
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "成功掃描 2 個圖面元素"
        
    def test_scan_elements_mixed_parameters(self):
        """混合參數測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {
                    "id": "AcDbCircle:1234567891",
                    "type": "circle",
                    "layer": "CIRCLES",
                    "color": 2
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"circle": 1},
                "layers": ["CIRCLES"],
                "bounds": {
                    "min": [10, 10, 0],
                    "max": [90, 90, 0]
                }
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.scan_elements(
            element_type="circle",
            include_geometry=True,
            include_properties=False,
            layer_filter="CIRCLES",
            bounds=[0, 0, 100, 100]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["scan_settings"]["element_type"] == "circle"
        assert result["data"]["scan_settings"]["include_geometry"] == True
        assert result["data"]["scan_settings"]["include_properties"] == False
        assert result["data"]["scan_settings"]["layer_filter"] == "CIRCLES"
        assert result["data"]["scan_settings"]["bounds"] == [0.0, 0.0, 100.0, 100.0]
        
    def test_scan_elements_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'scan_elements')
        assert callable(mcp_server_fastmcp.scan_elements)