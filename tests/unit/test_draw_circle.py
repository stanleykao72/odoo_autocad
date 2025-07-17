"""
TDD 測試：draw_circle MCP 工具

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


class TestDrawCircle:
    """draw_circle 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
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
        
    def test_draw_circle_with_custom_layer(self):
        """自定義圖層測試 - 應該失敗"""
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
        
    def test_draw_circle_parameter_validation_invalid_center_point_length(self):
        """參數驗證測試 - 無效圓心長度"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0],  # 只有2個元素
            radius=50
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_CENTER_POINT"
        assert "圓心座標必須是包含3個數值的列表" in result["message"]
        
    def test_draw_circle_parameter_validation_invalid_center_point_type(self):
        """參數驗證測試 - 無效圓心類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point="invalid",  # 不是列表
            radius=50
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_CENTER_POINT"
        
    def test_draw_circle_parameter_validation_non_numeric_center_coordinates(self):
        """參數驗證測試 - 非數值圓心座標"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, "z"],  # 非數值
            radius=50
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_CENTER_VALUES"
        assert "圓心座標必須是數值" in result["message"]
        
    def test_draw_circle_parameter_validation_invalid_radius_negative(self):
        """參數驗證測試 - 負半徑"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=-10  # 負數半徑
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_RADIUS"
        assert "半徑必須為正數" in result["message"]
        
    def test_draw_circle_parameter_validation_invalid_radius_zero(self):
        """參數驗證測試 - 零半徑"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=0  # 零半徑
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_RADIUS"
        
    def test_draw_circle_parameter_validation_invalid_radius_type(self):
        """參數驗證測試 - 無效半徑類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius="invalid"  # 非數值
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_RADIUS_TYPE"
        assert "半徑必須是數值" in result["message"]
        
    def test_draw_circle_parameter_validation_invalid_layer_empty(self):
        """參數驗證測試 - 無效圖層（空字串）"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50,
            layer=""  # 空字串
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        assert "圖層名稱必須是非空字串" in result["message"]
        
    def test_draw_circle_parameter_validation_invalid_layer_whitespace(self):
        """參數驗證測試 - 無效圖層（只有空白）"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50,
            layer="   "  # 只有空白
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_draw_circle_parameter_validation_invalid_layer_type(self):
        """參數驗證測試 - 無效圖層類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50,
            layer=123  # 不是字串
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_draw_circle_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_draw_circle_area_calculation_simple(self):
        """面積計算測試 - 簡單情況"""
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
        
    def test_draw_circle_area_calculation_decimal(self):
        """面積計算測試 - 小數半徑"""
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
            radius=5.5  # 半徑5.5，面積應為30.25π
        )
        
        # Assert
        assert result["status"] == "success"
        expected_area = math.pi * 5.5 * 5.5
        assert abs(result["data"]["area"] - expected_area) < 0.01
        
    def test_draw_circle_circumference_calculation_simple(self):
        """周長計算測試 - 簡單情況"""
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
        
    def test_draw_circle_circumference_calculation_decimal(self):
        """周長計算測試 - 小數半徑"""
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
            radius=2.5  # 半徑2.5，周長應為5π
        )
        
        # Assert
        assert result["status"] == "success"
        expected_circumference = 2 * math.pi * 2.5
        assert abs(result["data"]["circumference"] - expected_circumference) < 0.01
        
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
        
    def test_draw_circle_autocad_util_called_with_default_layer(self):
        """驗證使用預設圖層時的呼叫"""
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
        mock_autocad_util.draw_circle.assert_called_once_with(
            center_point=[0.0, 0.0, 0.0],
            radius=50.0,
            layer="0"
        )
        
    def test_draw_circle_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.side_effect = Exception("AutoCAD drawing error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "CIRCLE_DRAWING_ERROR"
        assert "繪製圓形失敗" in result["message"]
        
    def test_draw_circle_return_format_complete(self):
        """驗證回傳格式完整性"""
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
            layer="TEST_LAYER"
        )
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "circle_id" in data
        assert "center_point" in data
        assert "radius" in data
        assert "layer" in data
        assert "area" in data
        assert "circumference" in data
        assert "created_at" in data
        
        # 檢查時間戳格式
        assert isinstance(data["created_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_draw_circle_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
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
        assert result["data"]["created_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_draw_circle_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
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
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "draw_circle called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_draw_circle_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=50
        )
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in draw_circle" in call_args
        
    def test_draw_circle_coordinate_conversion(self):
        """座標轉換測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_circle.return_value = {
            "circle_id": "AcDbCircle:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[1, 2, 3],  # 整數
            radius=10.5  # 浮點數
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["center_point"] == [1.0, 2.0, 3.0]
        assert result["data"]["radius"] == 10.5
        
    def test_draw_circle_layer_strip(self):
        """圖層名稱去除空白測試"""
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
            layer="  TEST_LAYER  "  # 前後有空白
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer"] == "TEST_LAYER"
        mock_autocad_util.draw_circle.assert_called_once_with(
            center_point=[0.0, 0.0, 0.0],
            radius=50.0,
            layer="TEST_LAYER"
        )
        
    def test_draw_circle_radius_conversion(self):
        """半徑轉換測試"""
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
            radius=25  # 整數半徑
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["radius"] == 25.0
        
    def test_draw_circle_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'draw_circle')
        assert callable(mcp_server_fastmcp.draw_circle)