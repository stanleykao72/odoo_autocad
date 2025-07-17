"""
TDD 測試：list_layers MCP 工具

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


class TestListLayers:
    """list_layers 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
    def test_list_layers_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "0",
                    "color": 7,
                    "is_current": True,
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "on": True,
                    "frozen": False,
                    "locked": False,
                    "description": "Default layer"
                },
                {
                    "name": "WALLS",
                    "color": 2,
                    "is_current": False,
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "on": True,
                    "frozen": False,
                    "locked": False,
                    "description": "Wall elements"
                }
            ],
            "total_count": 2,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert len(result["data"]["layers"]) == 2
        assert result["data"]["total_count"] == 2
        assert result["data"]["current_layer"] == "0"
        assert result["data"]["filter_applied"] == "all"
        assert result["data"]["sort_applied"] == "name"
        assert "retrieved_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_list_layers_with_filter_visible(self):
        """過濾測試 - 可見圖層 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "0",
                    "color": 7,
                    "is_current": True,
                    "on": True,
                    "frozen": False,
                    "locked": False
                }
            ],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(filter_type="visible")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["filter_applied"] == "visible"
        assert result["data"]["total_count"] == 1
        
    def test_list_layers_parameter_validation_invalid_filter(self):
        """參數驗證測試 - 無效過濾類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(filter_type="invalid")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_FILTER_TYPE"
        assert "無效的過濾類型" in result["message"]
        
    def test_list_layers_parameter_validation_valid_filters(self):
        """參數驗證測試 - 有效過濾類型"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        valid_filters = ["all", "visible", "current", "frozen", "locked"]
        
        for filter_type in valid_filters:
            # Act
            result = mcp_server_fastmcp.list_layers(filter_type=filter_type)
            
            # Assert
            assert result["status"] == "success"
            assert result["data"]["filter_applied"] == filter_type
        
    def test_list_layers_parameter_validation_invalid_sort(self):
        """參數驗證測試 - 無效排序方式"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(sort_by="invalid")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_SORT_TYPE"
        assert "無效的排序方式" in result["message"]
        
    def test_list_layers_parameter_validation_valid_sorts(self):
        """參數驗證測試 - 有效排序方式"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        valid_sorts = ["name", "color", "created"]
        
        for sort_by in valid_sorts:
            # Act
            result = mcp_server_fastmcp.list_layers(sort_by=sort_by)
            
            # Assert
            assert result["status"] == "success"
            assert result["data"]["sort_applied"] == sort_by
        
    def test_list_layers_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_list_layers_with_sort_color(self):
        """排序測試 - 按顏色排序"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {"name": "WALLS", "color": 2},
                {"name": "0", "color": 7}
            ],
            "total_count": 2,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(sort_by="color")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sort_applied"] == "color"
        
    def test_list_layers_with_sort_created(self):
        """排序測試 - 按創建時間排序"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {"name": "0", "created": "2025-01-01"},
                {"name": "WALLS", "created": "2025-01-02"}
            ],
            "total_count": 2,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(sort_by="created")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sort_applied"] == "created"
        
    def test_list_layers_with_include_details_false(self):
        """詳細資訊測試 - 不包含詳細資訊"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {"name": "0", "color": 7},
                {"name": "WALLS", "color": 2}
            ],
            "total_count": 2,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(include_details=False)
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["total_count"] == 2
        
    def test_list_layers_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(
            filter_type="visible",
            sort_by="color",
            include_details=False
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.list_layers.assert_called_once_with(
            filter_type="visible",
            sort_by="color",
            include_details=False
        )
        
    def test_list_layers_autocad_util_called_with_defaults(self):
        """驗證使用預設值時的呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.list_layers.assert_called_once_with(
            filter_type="all",
            sort_by="name",
            include_details=True
        )
        
    def test_list_layers_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.side_effect = Exception("Layer listing error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "LAYER_LISTING_ERROR"
        assert "列出圖層失敗" in result["message"]
        
    def test_list_layers_empty_result(self):
        """空結果測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["total_count"] == 0
        assert len(result["data"]["layers"]) == 0
        assert result["message"] == "成功列出 0 個圖層"
        
    def test_list_layers_return_format_complete(self):
        """驗證回傳格式完整性"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "TEST_LAYER",
                    "color": 3,
                    "is_current": False,
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "on": True,
                    "frozen": False,
                    "locked": False
                }
            ],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "layers" in data
        assert "total_count" in data
        assert "current_layer" in data
        assert "filter_applied" in data
        assert "sort_applied" in data
        assert "retrieved_at" in data
        
        # 檢查時間戳格式
        assert isinstance(data["retrieved_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_list_layers_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["retrieved_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_list_layers_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "list_layers called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_list_layers_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in list_layers" in call_args
        
    def test_list_layers_filter_current(self):
        """過濾測試 - 當前圖層"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "0",
                    "color": 7,
                    "is_current": True
                }
            ],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(filter_type="current")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["filter_applied"] == "current"
        
    def test_list_layers_filter_frozen(self):
        """過濾測試 - 凍結圖層"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "FROZEN_LAYER",
                    "color": 7,
                    "frozen": True
                }
            ],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(filter_type="frozen")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["filter_applied"] == "frozen"
        
    def test_list_layers_filter_locked(self):
        """過濾測試 - 鎖定圖層"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "LOCKED_LAYER",
                    "color": 7,
                    "locked": True
                }
            ],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(filter_type="locked")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["filter_applied"] == "locked"
        
    def test_list_layers_multiple_layers(self):
        """多圖層測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {"name": "0", "color": 7, "is_current": True},
                {"name": "WALLS", "color": 2, "is_current": False},
                {"name": "DOORS", "color": 3, "is_current": False},
                {"name": "WINDOWS", "color": 4, "is_current": False}
            ],
            "total_count": 4,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["total_count"] == 4
        assert len(result["data"]["layers"]) == 4
        assert result["message"] == "成功列出 4 個圖層"
        
    def test_list_layers_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [{"name": "0"}],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "成功列出 1 個圖層"
        
    def test_list_layers_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'list_layers')
        assert callable(mcp_server_fastmcp.list_layers)