"""
TDD 測試：draw_line MCP 工具

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


class TestDrawLine:
    """draw_line 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
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
        assert "line_id" in result["data"]
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
    
    def test_draw_line_with_custom_layer(self):
        """自定義圖層測試 - 應該失敗"""
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
        
    def test_draw_line_parameter_validation_invalid_start_point_length(self):
        """參數驗證測試 - 無效起點長度"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0],  # 只有2個元素
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_START_POINT"
        assert "起點座標必須是包含3個數值的列表" in result["message"]
        
    def test_draw_line_parameter_validation_invalid_start_point_type(self):
        """參數驗證測試 - 無效起點類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point="invalid",  # 不是列表
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_START_POINT"
        
    def test_draw_line_parameter_validation_invalid_end_point_length(self):
        """參數驗證測試 - 無效終點長度"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0, 0]  # 4個元素
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_END_POINT"
        assert "終點座標必須是包含3個數值的列表" in result["message"]
        
    def test_draw_line_parameter_validation_invalid_end_point_type(self):
        """參數驗證測試 - 無效終點類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=None  # 不是列表
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_END_POINT"
        
    def test_draw_line_parameter_validation_non_numeric_start_coordinates(self):
        """參數驗證測試 - 非數值起點座標"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, "z"],  # 非數值
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COORDINATE_VALUES"
        assert "座標必須是數值" in result["message"]
        
    def test_draw_line_parameter_validation_non_numeric_end_coordinates(self):
        """參數驗證測試 - 非數值終點座標"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, "y", 0]  # 非數值
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COORDINATE_VALUES"
        
    def test_draw_line_parameter_validation_identical_points(self):
        """參數驗證測試 - 相同起終點"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[0, 0, 0]  # 相同點
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "IDENTICAL_POINTS"
        assert "起點和終點不能相同" in result["message"]
        
    def test_draw_line_parameter_validation_invalid_layer_empty(self):
        """參數驗證測試 - 無效圖層（空字串）"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0],
            layer=""  # 空字串
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        assert "圖層名稱必須是非空字串" in result["message"]
        
    def test_draw_line_parameter_validation_invalid_layer_whitespace(self):
        """參數驗證測試 - 無效圖層（只有空白）"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0],
            layer="   "  # 只有空白
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_draw_line_parameter_validation_invalid_layer_type(self):
        """參數驗證測試 - 無效圖層類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0],
            layer=123  # 不是字串
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_draw_line_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_draw_line_length_calculation_simple(self):
        """長度計算測試 - 簡單情況"""
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
        
    def test_draw_line_length_calculation_3d(self):
        """長度計算測試 - 3D 情況"""
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
            end_point=[1, 1, 1]  # 立方體對角線
        )
        
        # Assert
        assert result["status"] == "success"
        expected_length = math.sqrt(3)
        assert abs(result["data"]["length"] - expected_length) < 0.01
        
    def test_draw_line_angle_calculation_45_degrees(self):
        """角度計算測試 - 45度"""
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
        
    def test_draw_line_angle_calculation_90_degrees(self):
        """角度計算測試 - 90度"""
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
            end_point=[0, 1, 0]  # 90度角
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["angle"] == 90.0
        
    def test_draw_line_angle_calculation_negative(self):
        """角度計算測試 - 負角度"""
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
            end_point=[1, -1, 0]  # -45度角
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["angle"] == -45.0
        
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
        
    def test_draw_line_autocad_util_called_with_default_layer(self):
        """驗證使用預設圖層時的呼叫"""
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
        mock_autocad_util.draw_line.assert_called_once_with(
            start_point=[0.0, 0.0, 0.0],
            end_point=[100.0, 100.0, 0.0],
            layer="0"
        )
        
    def test_draw_line_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.side_effect = Exception("AutoCAD drawing error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "LINE_DRAWING_ERROR"
        assert "繪製直線失敗" in result["message"]
        
    def test_draw_line_return_format_complete(self):
        """驗證回傳格式完整性"""
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
            layer="TEST_LAYER"
        )
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "line_id" in data
        assert "start_point" in data
        assert "end_point" in data
        assert "layer" in data
        assert "length" in data
        assert "angle" in data
        assert "created_at" in data
        
        # 檢查時間戳格式
        assert isinstance(data["created_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_draw_line_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
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
        assert result["data"]["created_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_draw_line_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
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
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "draw_line called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_draw_line_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[100, 100, 0]
        )
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in draw_line" in call_args
        
    def test_draw_line_coordinate_conversion(self):
        """座標轉換測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.draw_line.return_value = {
            "line_id": "AcDbLine:1234567890",
            "success": True
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[1, 2, 3],  # 整數
            end_point=[4.5, 5.5, 6.5]  # 浮點數
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["start_point"] == [1.0, 2.0, 3.0]
        assert result["data"]["end_point"] == [4.5, 5.5, 6.5]
        
    def test_draw_line_layer_strip(self):
        """圖層名稱去除空白測試"""
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
            layer="  TEST_LAYER  "  # 前後有空白
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer"] == "TEST_LAYER"
        mock_autocad_util.draw_line.assert_called_once_with(
            start_point=[0.0, 0.0, 0.0],
            end_point=[100.0, 100.0, 0.0],
            layer="TEST_LAYER"
        )
        
    def test_draw_line_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'draw_line')
        assert callable(mcp_server_fastmcp.draw_line)