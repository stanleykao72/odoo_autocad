"""
TDD 測試：create_text MCP 工具

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


class TestCreateText:
    """create_text 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
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
        assert result["data"]["style"] == "Standard"
        assert result["data"]["alignment"] == "left"
        assert "properties" in result["data"]
        assert "bounds" in result["data"]
        assert "text_info" in result["data"]
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_create_text_with_custom_parameters(self):
        """自訂參數測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567891",
            "properties": {
                "color": 2,
                "linetype": "Continuous",
                "lineweight": "Default",
                "visible": True,
                "locked": False
            },
            "bounds": {
                "min": [200, 100, 0],
                "max": [240, 105, 0],
                "width": 40.0,
                "height": 5.0
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
            position=[200, 100, 0],
            text_content="標題文字",
            height=5.0,
            rotation=45.0,
            layer="TEXT",
            style="Arial",
            alignment="center"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["height"] == 5.0
        assert result["data"]["rotation"] == 45.0
        assert result["data"]["layer"] == "TEXT"
        assert result["data"]["style"] == "Arial"
        assert result["data"]["alignment"] == "center"
        
    def test_create_text_parameter_validation_invalid_position_length(self):
        """參數驗證測試 - 位置長度無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50],  # 只有2個座標
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POSITION"
        assert "文字位置格式無效" in result["message"]
        
    def test_create_text_parameter_validation_invalid_position_type(self):
        """參數驗證測試 - 位置類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position="not_a_list",
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POSITION"
        
    def test_create_text_parameter_validation_invalid_position_values(self):
        """參數驗證測試 - 位置值無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=["x", "y", "z"],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_POSITION_VALUES"
        assert "位置座標必須是數字" in result["message"]
        
    def test_create_text_parameter_validation_position_conversion(self):
        """參數驗證測試 - 位置座標轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=["100", "50", "0"],  # 字串數字
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["position"] == [100.0, 50.0, 0.0]
        
    def test_create_text_parameter_validation_empty_text_content(self):
        """參數驗證測試 - 文字內容為空"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content=""
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "EMPTY_TEXT_CONTENT"
        assert "文字內容不能為空" in result["message"]
        
    def test_create_text_parameter_validation_whitespace_text_content(self):
        """參數驗證測試 - 文字內容只有空白"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="   "
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "EMPTY_TEXT_CONTENT"
        
    def test_create_text_parameter_validation_text_content_strip(self):
        """參數驗證測試 - 文字內容去除空白"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="  測試文字  "
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["text_content"] == "測試文字"
        
    def test_create_text_parameter_validation_invalid_height_negative(self):
        """參數驗證測試 - 高度為負數"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height=-1.0
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_HEIGHT"
        assert "文字高度必須大於 0" in result["message"]
        
    def test_create_text_parameter_validation_invalid_height_zero(self):
        """參數驗證測試 - 高度為零"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height=0.0
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_HEIGHT"
        
    def test_create_text_parameter_validation_invalid_height_type(self):
        """參數驗證測試 - 高度類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height="not_a_number"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_HEIGHT_TYPE"
        assert "文字高度必須是數字" in result["message"]
        
    def test_create_text_parameter_validation_valid_height_conversion(self):
        """參數驗證測試 - 高度轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height="5.0"  # 字串數字
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["height"] == 5.0
        
    def test_create_text_parameter_validation_invalid_rotation_type(self):
        """參數驗證測試 - 旋轉角度類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            rotation="not_a_number"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ROTATION_TYPE"
        assert "旋轉角度必須是數字" in result["message"]
        
    def test_create_text_parameter_validation_valid_rotation_conversion(self):
        """參數驗證測試 - 旋轉角度轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            rotation="45.0"  # 字串數字
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["rotation"] == 45.0
        
    def test_create_text_parameter_validation_invalid_layer_empty(self):
        """參數驗證測試 - 圖層名稱為空"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            layer=""
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER"
        assert "圖層名稱不能為空" in result["message"]
        
    def test_create_text_parameter_validation_invalid_layer_whitespace(self):
        """參數驗證測試 - 圖層名稱只有空白"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            layer="   "
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER"
        
    def test_create_text_parameter_validation_layer_strip(self):
        """參數驗證測試 - 圖層名稱去除空白"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            layer="  TEXT  "
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer"] == "TEXT"
        
    def test_create_text_parameter_validation_invalid_alignment(self):
        """參數驗證測試 - 對齊方式無效"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            alignment="invalid"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ALIGNMENT"
        assert "無效的對齊方式" in result["message"]
        
    def test_create_text_parameter_validation_valid_alignments(self):
        """參數驗證測試 - 有效對齊方式"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        valid_alignments = ["left", "center", "right"]
        
        for alignment in valid_alignments:
            # Act
            result = mcp_server_fastmcp.create_text(
                position=[100, 50, 0],
                text_content="測試文字",
                alignment=alignment
            )
            
            # Assert
            assert result["status"] == "success"
            assert result["data"]["alignment"] == alignment
        
    def test_create_text_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_create_text_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            height=5.0,
            rotation=45.0,
            layer="TEXT",
            style="Arial",
            alignment="center"
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.create_text.assert_called_once()
        call_args = mock_autocad_util.create_text.call_args
        assert call_args.kwargs["position"] == [100.0, 50.0, 0.0]
        assert call_args.kwargs["text_content"] == "測試文字"
        assert call_args.kwargs["height"] == 5.0
        assert call_args.kwargs["rotation"] == math.radians(45.0)
        assert call_args.kwargs["layer"] == "TEXT"
        assert call_args.kwargs["style"] == "Arial"
        assert call_args.kwargs["alignment"] == "center"
        
    def test_create_text_autocad_util_called_with_defaults(self):
        """驗證使用預設值時的呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.create_text.assert_called_once()
        call_args = mock_autocad_util.create_text.call_args
        assert call_args.kwargs["position"] == [100.0, 50.0, 0.0]
        assert call_args.kwargs["text_content"] == "測試文字"
        assert call_args.kwargs["height"] == 2.5
        assert call_args.kwargs["rotation"] == 0.0
        assert call_args.kwargs["layer"] == "0"
        assert call_args.kwargs["style"] == "Standard"
        assert call_args.kwargs["alignment"] == "left"
        
    def test_create_text_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.side_effect = Exception("Text creation error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "TEXT_CREATION_ERROR"
        assert "創建文字失敗" in result["message"]
        
    def test_create_text_return_format_complete(self):
        """驗證回傳格式完整性"""
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
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "text_id" in data
        assert "position" in data
        assert "text_content" in data
        assert "height" in data
        assert "rotation" in data
        assert "layer" in data
        assert "style" in data
        assert "alignment" in data
        assert "properties" in data
        assert "bounds" in data
        assert "text_info" in data
        assert "created_at" in data
        
        # 檢查 properties 結構
        properties = data["properties"]
        assert "color" in properties
        assert "linetype" in properties
        assert "lineweight" in properties
        assert "visible" in properties
        assert "locked" in properties
        
        # 檢查 bounds 結構
        bounds = data["bounds"]
        assert "min" in bounds
        assert "max" in bounds
        assert "width" in bounds
        assert "height" in bounds
        
        # 檢查 text_info 結構
        text_info = data["text_info"]
        assert "character_count" in text_info
        assert "line_count" in text_info
        assert "font_name" in text_info
        assert "is_bold" in text_info
        assert "is_italic" in text_info
        
        # 檢查時間戳格式
        assert isinstance(data["created_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_create_text_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["created_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_create_text_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "create_text called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_create_text_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in create_text" in call_args
        
    def test_create_text_rotation_conversion_to_radians(self):
        """旋轉角度轉換為弧度測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字",
            rotation=90.0
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.create_text.assert_called_once()
        call_args = mock_autocad_util.create_text.call_args
        assert abs(call_args.kwargs["rotation"] - math.radians(90.0)) < 0.001
        
    def test_create_text_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "成功創建文字註解: 測試文字"
        
    def test_create_text_default_values(self):
        """預設值測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="測試文字"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["height"] == 2.5
        assert result["data"]["rotation"] == 0.0
        assert result["data"]["layer"] == "0"
        assert result["data"]["style"] == "Standard"
        assert result["data"]["alignment"] == "left"
        
    def test_create_text_multiline_text(self):
        """多行文字測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_text.return_value = {
            "text_id": "AcDbText:1234567890",
            "properties": {},
            "bounds": {},
            "text_info": {
                "character_count": 11,
                "line_count": 2,
                "font_name": "Arial",
                "is_bold": False,
                "is_italic": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_text(
            position=[100, 50, 0],
            text_content="第一行\n第二行"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["text_content"] == "第一行\n第二行"
        
    def test_create_text_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'create_text')
        assert callable(mcp_server_fastmcp.create_text)