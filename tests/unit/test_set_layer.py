"""
TDD 測試：set_layer MCP 工具

遵循 Red-Green-Refactor 循環：
1. Red: 寫失敗測試 (此階段)
2. Green: 最小實現
3. Refactor: 完整實現

測試覆蓋率要求: 95%+
"""

import pytest
from unittest.mock import patch, Mock, MagicMock
from datetime import datetime
import os
import sys

# 添加專案根目錄到 Python 路徑
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import mcp_server_fastmcp


class TestSetLayer:
    """set_layer 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
    def test_set_layer_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "WALLS",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {
                "name": "WALLS",
                "color": 7,
                "linetype": "Continuous",
                "lineweight": "Default",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="WALLS")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer_name"] == "WALLS"
        assert result["data"]["color"] == 7
        assert result["data"]["is_current"] == True
        assert result["data"]["created"] == False
        assert "layer_info" in result["data"]
        assert "changed_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_set_layer_with_custom_color(self):
        """自定義顏色測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "DIMENSIONS",
            "color": 2,
            "is_current": True,
            "created": True,
            "layer_info": {
                "name": "DIMENSIONS",
                "color": 2,
                "linetype": "Continuous",
                "lineweight": "Default",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="DIMENSIONS",
            color=2
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer_name"] == "DIMENSIONS"
        assert result["data"]["color"] == 2
        assert result["data"]["created"] == True
        
    def test_set_layer_parameter_validation_empty_name(self):
        """參數驗證測試 - 空白圖層名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        assert "圖層名稱不能為空" in result["message"]
        
    def test_set_layer_parameter_validation_whitespace_name(self):
        """參數驗證測試 - 只有空白的圖層名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="   ")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_set_layer_parameter_validation_invalid_layer_type(self):
        """參數驗證測試 - 無效圖層類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name=123)
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_LAYER_NAME"
        
    def test_set_layer_parameter_validation_invalid_color_negative(self):
        """參數驗證測試 - 負數顏色"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST",
            color=-1
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COLOR"
        assert "顏色必須是 0-255 之間的整數" in result["message"]
        
    def test_set_layer_parameter_validation_invalid_color_high(self):
        """參數驗證測試 - 過高顏色值"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST",
            color=256
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COLOR"
        
    def test_set_layer_parameter_validation_invalid_color_type(self):
        """參數驗證測試 - 無效顏色類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST",
            color="red"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_COLOR"
        
    def test_set_layer_parameter_validation_valid_color_boundaries(self):
        """參數驗證測試 - 有效顏色邊界值"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TEST",
            "color": 0,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert - 測試 0
        result = mcp_server_fastmcp.set_layer(layer_name="TEST", color=0)
        assert result["status"] == "success"
        assert result["data"]["color"] == 0
        
        # Act & Assert - 測試 255
        mock_autocad_util.set_layer.return_value["color"] = 255
        result = mcp_server_fastmcp.set_layer(layer_name="TEST", color=255)
        assert result["status"] == "success"
        assert result["data"]["color"] == 255
        
    def test_set_layer_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="TEST")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_set_layer_create_if_not_exist_true(self):
        """自動創建圖層測試 - True"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "NEW_LAYER",
            "color": 7,
            "is_current": True,
            "created": True,
            "layer_info": {
                "name": "NEW_LAYER",
                "color": 7,
                "linetype": "Continuous",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="NEW_LAYER",
            create_if_not_exist=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["created"] == True
        
    def test_set_layer_create_if_not_exist_false(self):
        """不自動創建圖層測試 - False"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "EXISTING",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {
                "name": "EXISTING",
                "color": 7,
                "linetype": "Continuous",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="EXISTING",
            create_if_not_exist=False
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["created"] == False
        
    def test_set_layer_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TEST_LAYER",
            "color": 5,
            "is_current": True,
            "created": True,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="TEST_LAYER",
            color=5,
            create_if_not_exist=True
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="TEST_LAYER",
            color=5,
            create_if_not_exist=True
        )
        
    def test_set_layer_autocad_util_called_with_defaults(self):
        """驗證使用預設值時的呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "DEFAULT_TEST",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="DEFAULT_TEST")
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="DEFAULT_TEST",
            color=7,
            create_if_not_exist=True
        )
        
    def test_set_layer_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.side_effect = Exception("Layer setting error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="TEST")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "LAYER_SETTING_ERROR"
        assert "設定圖層失敗" in result["message"]
        
    def test_set_layer_layer_name_strip(self):
        """圖層名稱去除空白測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TRIMMED",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="  TRIMMED  ")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer_name"] == "TRIMMED"
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="TRIMMED",
            color=7,
            create_if_not_exist=True
        )
        
    def test_set_layer_return_format_complete(self):
        """驗證回傳格式完整性"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "COMPLETE_TEST",
            "color": 3,
            "is_current": True,
            "created": True,
            "layer_info": {
                "name": "COMPLETE_TEST",
                "color": 3,
                "linetype": "Continuous",
                "lineweight": "Default",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="COMPLETE_TEST",
            color=3
        )
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "layer_name" in data
        assert "color" in data
        assert "is_current" in data
        assert "created" in data
        assert "layer_info" in data
        assert "changed_at" in data
        
        # 檢查時間戳格式
        assert isinstance(data["changed_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_set_layer_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TIMESTAMP_TEST",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="TIMESTAMP_TEST")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["changed_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_set_layer_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "LOG_TEST",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="LOG_TEST")
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "set_layer called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_set_layer_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="ERROR_TEST")
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in set_layer" in call_args
        
    def test_set_layer_create_if_not_exist_type_validation(self):
        """create_if_not_exist 參數類型驗證"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "TYPE_TEST",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act - 測試布林值
        result = mcp_server_fastmcp.set_layer(
            layer_name="TYPE_TEST",
            create_if_not_exist=False
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="TYPE_TEST",
            color=7,
            create_if_not_exist=False
        )
        
    def test_set_layer_multiple_parameters(self):
        """多參數組合測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "MULTI_TEST",
            "color": 4,
            "is_current": True,
            "created": True,
            "layer_info": {
                "name": "MULTI_TEST",
                "color": 4,
                "linetype": "Continuous",
                "on": True,
                "frozen": False,
                "locked": False
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(
            layer_name="MULTI_TEST",
            color=4,
            create_if_not_exist=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["layer_name"] == "MULTI_TEST"
        assert result["data"]["color"] == 4
        assert result["data"]["created"] == True
        mock_autocad_util.set_layer.assert_called_once_with(
            layer_name="MULTI_TEST",
            color=4,
            create_if_not_exist=True
        )
        
    def test_set_layer_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.set_layer.return_value = {
            "layer_name": "MESSAGE_TEST",
            "color": 7,
            "is_current": True,
            "created": False,
            "layer_info": {}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.set_layer(layer_name="MESSAGE_TEST")
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "成功設定圖層: MESSAGE_TEST"
        
    def test_set_layer_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'set_layer')
        assert callable(mcp_server_fastmcp.set_layer)