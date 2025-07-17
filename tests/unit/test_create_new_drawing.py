"""
TDD 測試：create_new_drawing MCP 工具

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


class TestCreateNewDrawing:
    """create_new_drawing 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
    def test_create_new_drawing_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "新圖面.dwg",
            "full_path": "C:\\Projects\\新圖面.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "新圖面.dwg"
        assert result["data"]["full_path"] == "C:\\Projects\\新圖面.dwg"
        assert result["data"]["template_used"] == "預設模板"
        assert result["data"]["units"] == "公制"
        assert "created_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_create_new_drawing_with_custom_parameters(self):
        """自定義參數測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "工程圖_A1.dwg",
            "full_path": "C:\\Projects\\工程圖_A1.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        with patch('mcp_server_fastmcp.os.path.exists', return_value=True):
            result = mcp_server_fastmcp.create_new_drawing(
                drawing_name="工程圖_A1",
                units="英制",
                template_path="C:\\Templates\\standard.dwt",
                save_path="C:\\Projects\\工程圖_A1.dwg"
            )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "工程圖_A1.dwg"
        assert result["data"]["units"] == "英制"
        assert result["data"]["template_used"] == "C:\\Templates\\standard.dwt"
        
    def test_create_new_drawing_parameter_validation_empty_name(self):
        """參數驗證測試 - 空白圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(drawing_name="")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        assert "圖面名稱不能為空" in result["message"]
        
    def test_create_new_drawing_parameter_validation_whitespace_name(self):
        """參數驗證測試 - 只有空白字元的圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(drawing_name="   ")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        
    def test_create_new_drawing_parameter_validation_invalid_units(self):
        """參數驗證測試 - 無效單位"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(units="無效單位")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_UNITS"
        assert "單位系統必須是 '公制' 或 '英制'" in result["message"]
        
    def test_create_new_drawing_valid_units_metric(self):
        """參數驗證測試 - 有效單位（公制）"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(units="公制")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["units"] == "公制"
        
    def test_create_new_drawing_valid_units_imperial(self):
        """參數驗證測試 - 有效單位（英制）"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(units="英制")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["units"] == "英制"
        
    def test_create_new_drawing_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing()
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_create_new_drawing_template_not_found(self):
        """模板檔案不存在測試"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        with patch('mcp_server_fastmcp.os.path.exists', return_value=False):
            result = mcp_server_fastmcp.create_new_drawing(template_path="non_existent.dwt")
            
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "TEMPLATE_NOT_FOUND"
        assert "模板檔案不存在" in result["message"]
        
    def test_create_new_drawing_template_exists(self):
        """模板檔案存在測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        with patch('mcp_server_fastmcp.os.path.exists', return_value=True):
            result = mcp_server_fastmcp.create_new_drawing(template_path="existing.dwt")
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["template_used"] == "existing.dwt"
        
    def test_create_new_drawing_autocad_util_exception(self):
        """AutoCAD 工具異常測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.side_effect = Exception("AutoCAD error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing()
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "DRAWING_CREATION_ERROR"
        assert "創建圖面失敗" in result["message"]
        
    def test_create_new_drawing_autocad_util_called_with_correct_params(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        with patch('mcp_server_fastmcp.os.path.exists', return_value=True):
            result = mcp_server_fastmcp.create_new_drawing(
                drawing_name="測試圖面",
                template_path="template.dwt",
                units="英制",
                save_path="C:\\Projects\\測試.dwg"
            )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.create_new_drawing.assert_called_once_with(
            drawing_name="測試圖面",
            template_path="template.dwt",
            units="英制",
            save_path="C:\\Projects\\測試.dwg"
        )
        
    def test_create_new_drawing_autocad_util_called_with_none_defaults(self):
        """驗證空字串參數被轉換為 None"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(
            drawing_name="測試圖面",
            template_path="",
            save_path=""
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.create_new_drawing.assert_called_once_with(
            drawing_name="測試圖面",
            template_path=None,
            units="公制",
            save_path=None
        )
        
    def test_create_new_drawing_return_format_complete(self):
        """驗證回傳格式完整性"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "完整測試.dwg",
            "full_path": "C:\\Projects\\完整測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(drawing_name="完整測試")
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "drawing_name" in data
        assert "full_path" in data
        assert "template_used" in data
        assert "units" in data
        assert "created_at" in data
        
        # 檢查時間戳格式
        assert isinstance(data["created_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_create_new_drawing_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["created_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_create_new_drawing_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.return_value = {
            "drawing_name": "測試.dwg",
            "full_path": "C:\\Projects\\測試.dwg"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "create_new_drawing called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_create_new_drawing_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.create_new_drawing.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.create_new_drawing()
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in create_new_drawing" in call_args
        
    def test_create_new_drawing_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'create_new_drawing')
        assert callable(mcp_server_fastmcp.create_new_drawing)