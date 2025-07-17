"""
TDD 測試：add_dimension MCP 工具

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


class TestAddDimension:
    """add_dimension 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
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
        assert result["data"]["dim_style"] == "Standard"
        assert result["data"]["layer"] == "0"
        assert result["data"]["angle"] == 0.0
        assert result["data"]["text_override"] is None
        assert "properties" in result["data"]
        assert "dimension_info" in result["data"]
        assert "bounds" in result["data"]
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_add_dimension_angular_functionality(self):
        """角度尺寸功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567891",
            "measured_value": 45.0,
            "display_text": "45°",
            "properties": {"color": 7},
            "dimension_info": {"units": "degrees"},
            "bounds": {"min": [0, 0, 0], "max": [50, 50, 0]}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="angular",
            definition_points=[[0, 0, 0], [50, 0, 0], [50, 50, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["dimension_type"] == "angular"
        assert result["data"]["measured_value"] == 45.0
        assert result["data"]["display_text"] == "45°"
        
    def test_add_dimension_radial_functionality(self):
        """徑向尺寸功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567892",
            "measured_value": 25.0,
            "display_text": "R25.00",
            "properties": {"color": 7},
            "dimension_info": {"units": "mm"},
            "bounds": {"min": [0, 0, 0], "max": [50, 25, 0]}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="radial",
            definition_points=[[0, 0, 0], [25, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["dimension_type"] == "radial"
        assert result["data"]["measured_value"] == 25.0
        assert result["data"]["display_text"] == "R25.00"
        
    def test_add_dimension_diameter_functionality(self):
        """直徑尺寸功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567893",
            "measured_value": 50.0,
            "display_text": "Ø50.00",
            "properties": {"color": 7},
            "dimension_info": {"units": "mm"},
            "bounds": {"min": [-25, -25, 0], "max": [25, 25, 0]}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="diameter",
            definition_points=[[-25, 0, 0], [25, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["dimension_type"] == "diameter"
        assert result["data"]["measured_value"] == 50.0
        assert result["data"]["display_text"] == "Ø50.00"
        
    def test_add_dimension_with_custom_parameters(self):
        """自訂參數測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567894",
            "measured_value": 150.0,
            "display_text": "Custom 150",
            "properties": {"color": 2},
            "dimension_info": {"units": "mm"},
            "bounds": {"min": [0, 0, 0], "max": [150, 20, 0]}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [150, 0, 0]],
            text_position=[75, 10, 0],
            text_override="Custom 150",
            dim_style="ISO-25",
            layer="DIMENSIONS",
            angle=30.0
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["text_position"] == [75.0, 10.0, 0.0]
        assert result["data"]["text_override"] == "Custom 150"
        assert result["data"]["dim_style"] == "ISO-25"
        assert result["data"]["layer"] == "DIMENSIONS"
        assert result["data"]["angle"] == 30.0
        
    def test_add_dimension_parameter_validation_invalid_type(self):
        """參數驗證測試 - 無效尺寸類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="invalid_type",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DIMENSION_TYPE"
        assert "無效的尺寸類型" in result["message"]
        
    def test_add_dimension_parameter_validation_valid_types(self):
        """參數驗證測試 - 有效尺寸類型"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        valid_types = ["linear", "angular", "radial", "diameter"]
        
        for dim_type in valid_types:
            # Act
            result = mcp_server_fastmcp.add_dimension(
                dimension_type=dim_type,
                definition_points=[[0, 0, 0], [100, 0, 0]]
            )
            
            # Assert
            assert result["status"] == "success"
            assert result["data"]["dimension_type"] == dim_type
        
    def test_add_dimension_parameter_validation_insufficient_points(self):
        """參數驗證測試 - 定義點不足"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0]]  # 只有一個點
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DEFINITION_POINTS"
        assert "定義點必須至少包含兩個點" in result["message"]
        
    def test_add_dimension_parameter_validation_invalid_definition_points_not_list(self):
        """參數驗證測試 - 定義點不是列表"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points="not_a_list"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DEFINITION_POINTS"
        
    def test_add_dimension_parameter_validation_invalid_point_format(self):
        """參數驗證測試 - 定義點格式無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0], [100, 0]]  # 缺少z座標
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POINT_FORMAT"
        assert "定義點 1 格式無效" in result["message"]
        
    def test_add_dimension_parameter_validation_invalid_point_values(self):
        """參數驗證測試 - 定義點值無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[["x", "y", "z"], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POINT_VALUES"
        assert "定義點 1 座標必須是數字" in result["message"]
        
    def test_add_dimension_parameter_validation_point_conversion(self):
        """參數驗證測試 - 定義點座標轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[["0", "0", "0"], ["100", "0", "0"]]  # 字串數字
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["definition_points"] == [[0.0, 0.0, 0.0], [100.0, 0.0, 0.0]]
        
    def test_add_dimension_parameter_validation_invalid_text_position_format(self):
        """參數驗證測試 - 文字位置格式無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            text_position=[50, 10]  # 缺少z座標
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_TEXT_POSITION"
        assert "文字位置格式無效" in result["message"]
        
    def test_add_dimension_parameter_validation_invalid_text_position_values(self):
        """參數驗證測試 - 文字位置值無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            text_position=["x", "y", "z"]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_TEXT_POSITION_VALUES"
        assert "文字位置座標必須是數字" in result["message"]
        
    def test_add_dimension_parameter_validation_text_position_conversion(self):
        """參數驗證測試 - 文字位置座標轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            text_position=["50", "10", "0"]  # 字串數字
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["text_position"] == [50.0, 10.0, 0.0]
        
    def test_add_dimension_parameter_validation_invalid_angle_type(self):
        """參數驗證測試 - 角度類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            angle="not_a_number"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ANGLE_TYPE"
        assert "角度必須是數字" in result["message"]
        
    def test_add_dimension_parameter_validation_valid_angle_conversion(self):
        """參數驗證測試 - 角度轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            angle="45.0"  # 字串數字
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["angle"] == 45.0
        
    def test_add_dimension_parameter_validation_invalid_layer_empty(self):
        """參數驗證測試 - 圖層名稱為空"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            layer=""
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER"
        assert "圖層名稱不能為空" in result["message"]
        
    def test_add_dimension_parameter_validation_invalid_layer_whitespace(self):
        """參數驗證測試 - 圖層名稱只有空白"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            layer="   "
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER"
        
    def test_add_dimension_parameter_validation_layer_strip(self):
        """參數驗證測試 - 圖層名稱去除空白"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            layer="  DIMENSIONS  "
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer"] == "DIMENSIONS"
        
    def test_add_dimension_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_add_dimension_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            text_position=[50, 10, 0],
            text_override="Custom Text",
            dim_style="ISO-25",
            layer="DIMENSIONS",
            angle=30.0
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.add_dimension.assert_called_once()
        call_args = mock_autocad_util.add_dimension.call_args
        assert call_args.kwargs["dimension_type"] == "linear"
        assert call_args.kwargs["definition_points"] == [[0.0, 0.0, 0.0], [100.0, 0.0, 0.0]]
        assert call_args.kwargs["text_position"] == [50.0, 10.0, 0.0]
        assert call_args.kwargs["text_override"] == "Custom Text"
        assert call_args.kwargs["dim_style"] == "ISO-25"
        assert call_args.kwargs["layer"] == "DIMENSIONS"
        assert abs(call_args.kwargs["angle"] - math.radians(30.0)) < 0.001
        
    def test_add_dimension_autocad_util_called_with_defaults(self):
        """驗證使用預設值時的呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.add_dimension.assert_called_once()
        call_args = mock_autocad_util.add_dimension.call_args
        assert call_args.kwargs["dimension_type"] == "linear"
        assert call_args.kwargs["definition_points"] == [[0.0, 0.0, 0.0], [100.0, 0.0, 0.0]]
        assert call_args.kwargs["text_position"] is None
        assert call_args.kwargs["text_override"] is None
        assert call_args.kwargs["dim_style"] == "Standard"
        assert call_args.kwargs["layer"] == "0"
        assert call_args.kwargs["angle"] == 0.0
        
    def test_add_dimension_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.side_effect = Exception("Dimension creation error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "DIMENSION_CREATION_ERROR"
        assert "添加尺寸標註失敗" in result["message"]
        
    def test_add_dimension_return_format_complete(self):
        """驗證回傳格式完整性"""
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
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "dimension_id" in data
        assert "dimension_type" in data
        assert "definition_points" in data
        assert "text_position" in data
        assert "measured_value" in data
        assert "display_text" in data
        assert "text_override" in data
        assert "dim_style" in data
        assert "layer" in data
        assert "angle" in data
        assert "properties" in data
        assert "dimension_info" in data
        assert "bounds" in data
        assert "created_at" in data
        
        # 檢查 properties 結構
        properties = data["properties"]
        assert "color" in properties
        assert "linetype" in properties
        assert "lineweight" in properties
        assert "visible" in properties
        assert "locked" in properties
        
        # 檢查 dimension_info 結構
        dimension_info = data["dimension_info"]
        assert "units" in dimension_info
        assert "precision" in dimension_info
        assert "scale" in dimension_info
        assert "arrow_size" in dimension_info
        assert "text_height" in dimension_info
        
        # 檢查 bounds 結構
        bounds = data["bounds"]
        assert "min" in bounds
        assert "max" in bounds
        assert "width" in bounds
        assert "height" in bounds
        
        # 檢查時間戳格式
        assert isinstance(data["created_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_add_dimension_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["created_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_add_dimension_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "add_dimension called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_add_dimension_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in add_dimension" in call_args
        
    def test_add_dimension_angle_conversion_to_radians(self):
        """角度轉換為弧度測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]],
            angle=90.0
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.add_dimension.assert_called_once()
        call_args = mock_autocad_util.add_dimension.call_args
        assert abs(call_args.kwargs["angle"] - math.radians(90.0)) < 0.001
        
    def test_add_dimension_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "成功添加linear尺寸標註: 100.00"
        
    def test_add_dimension_default_values(self):
        """預設值測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.add_dimension.return_value = {
            "dimension_id": "AcDbDimension:1234567890",
            "measured_value": 100.0,
            "display_text": "100.00",
            "properties": {},
            "dimension_info": {},
            "bounds": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.add_dimension(
            dimension_type="linear",
            definition_points=[[0, 0, 0], [100, 0, 0]]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["text_position"] is None
        assert result["data"]["text_override"] is None
        assert result["data"]["dim_style"] == "Standard"
        assert result["data"]["layer"] == "0"
        assert result["data"]["angle"] == 0.0
        
    def test_add_dimension_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'add_dimension')
        assert callable(mcp_server_fastmcp.add_dimension)