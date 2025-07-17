"""
TDD 測試：export_to_database MCP 工具

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
import tempfile
import sqlite3

# 添加專案根目錄到 Python 路徑
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import mcp_server_fastmcp


class TestExportToDatabase:
    """export_to_database 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
    def test_export_to_database_basic_functionality(self):
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
                        "lineweight": "Default"
                    },
                    "bounds": {
                        "min": [0, 0, 0],
                        "max": [100, 100, 0]
                    }
                }
            ],
            "summary": {
                "total_count": 1,
                "element_counts": {"line": 1}
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "測試圖面"
        assert result["data"]["export_summary"]["total_elements"] == 1
        assert result["data"]["export_summary"]["element_breakdown"]["line"] == 1
        assert "database_path" in result["data"]
        assert "exported_at" in result["data"]
        assert "performance" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_export_to_database_parameter_validation_empty_name(self):
        """參數驗證測試 - 空白圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        assert "圖面名稱不能為空" in result["message"]
        
    def test_export_to_database_parameter_validation_whitespace_name(self):
        """參數驗證測試 - 只有空白的圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="   ")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        
    def test_export_to_database_parameter_validation_invalid_element_types_not_list(self):
        """參數驗證測試 - 元素類型不是列表"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types="not_a_list"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ELEMENT_TYPES"
        assert "元素類型必須是列表格式" in result["message"]
        
    def test_export_to_database_parameter_validation_invalid_element_types_values(self):
        """參數驗證測試 - 無效元素類型值"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types=["invalid_type", "another_invalid"]
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_ELEMENT_TYPES"
        assert "無效的元素類型" in result["message"]
        
    def test_export_to_database_parameter_validation_valid_element_types(self):
        """參數驗證測試 - 有效元素類型"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        valid_types = ["line", "circle", "arc", "text", "dimension", "block"]
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types=valid_types
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["element_types"] == valid_types
        
    def test_export_to_database_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_export_to_database_with_specific_element_types(self):
        """特定元素類型匯出測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {"id": "AcDbLine:1", "type": "line", "layer": "0", "color": 7},
                {"id": "AcDbCircle:2", "type": "circle", "layer": "0", "color": 7}
            ],
            "summary": {"total_count": 2}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types=["line", "circle"]
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["element_types"] == ["line", "circle"]
        
    def test_export_to_database_incremental_update_false(self):
        """非增量更新測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            incremental_update=False
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["incremental_update"] == False
        
    def test_export_to_database_with_layer_filter(self):
        """圖層過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            layer_filter="WALLS"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["layer_filter"] == "WALLS"
        
    def test_export_to_database_without_geometry(self):
        """不包含幾何資料測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            include_geometry=False
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["include_geometry"] == False
        
    def test_export_to_database_with_odoo_sync_success(self):
        """Odoo 同步成功測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_elements.return_value = {"success": True}
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            sync_to_odoo=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_info"]["odoo_sync"] == True
        assert result["data"]["sync_info"]["sync_status"] == "success"
        
    def test_export_to_database_with_odoo_sync_failed(self):
        """Odoo 同步失敗測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_elements.return_value = {"success": False}
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            sync_to_odoo=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_info"]["odoo_sync"] == True
        assert result["data"]["sync_info"]["sync_status"] == "failed"
        
    def test_export_to_database_with_odoo_sync_no_connection(self):
        """Odoo 同步無連接測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        mcp_server_fastmcp._odoo_util = None
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            sync_to_odoo=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_info"]["odoo_sync"] == True
        assert result["data"]["sync_info"]["sync_status"] == "failed"
        assert "Odoo 連接未建立" in result["data"]["sync_info"]["error"]
        
    def test_export_to_database_with_odoo_sync_exception(self):
        """Odoo 同步異常測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_elements.side_effect = Exception("Sync error")
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            sync_to_odoo=True
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_info"]["odoo_sync"] == True
        assert result["data"]["sync_info"]["sync_status"] == "failed"
        assert "Sync error" in result["data"]["sync_info"]["error"]
        
    def test_export_to_database_multiple_elements(self):
        """多元素匯出測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {"id": "AcDbLine:1", "type": "line", "layer": "0", "color": 7},
                {"id": "AcDbCircle:2", "type": "circle", "layer": "CIRCLES", "color": 2},
                {"id": "AcDbText:3", "type": "text", "layer": "TEXT", "color": 3}
            ],
            "summary": {
                "total_count": 3,
                "element_counts": {"line": 1, "circle": 1, "text": 1}
            }
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_summary"]["total_elements"] == 3
        assert result["data"]["export_summary"]["element_breakdown"]["line"] == 1
        assert result["data"]["export_summary"]["element_breakdown"]["circle"] == 1
        assert result["data"]["export_summary"]["element_breakdown"]["text"] == 1
        
    def test_export_to_database_scan_elements_called_correctly(self):
        """驗證 scan_elements 被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            include_geometry=False,
            layer_filter="WALLS"
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.scan_elements.assert_called_once_with(
            element_type="all",
            include_geometry=False,
            include_properties=True,
            layer_filter="WALLS"
        )
        
    def test_export_to_database_scan_elements_called_with_element_types(self):
        """驗證指定元素類型時的呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="測試圖面",
            element_types=["line", "circle"]
        )
        
        # Assert
        assert result["status"] == "success"
        # 應該為每個元素類型呼叫一次
        assert mock_autocad_util.scan_elements.call_count == 2
        
    def test_export_to_database_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.side_effect = Exception("Scan error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "DATABASE_EXPORT_ERROR"
        assert "匯出到資料庫失敗" in result["message"]
        
    def test_export_to_database_empty_elements(self):
        """空元素測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0, "element_counts": {}}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_summary"]["total_elements"] == 0
        assert result["data"]["export_summary"]["element_breakdown"] == {}
        assert result["message"] == "成功匯出 0 個元素到資料庫 (新增: 0, 更新: 0)"
        
    def test_export_to_database_return_format_complete(self):
        """驗證回傳格式完整性"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {"id": "AcDbLine:1", "type": "line", "layer": "0", "color": 7}
            ],
            "summary": {"total_count": 1, "element_counts": {"line": 1}}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "drawing_name" in data
        assert "database_path" in data
        assert "export_summary" in data
        assert "database_info" in data
        assert "sync_info" in data
        assert "export_settings" in data
        assert "performance" in data
        assert "exported_at" in data
        
        # 檢查 export_summary 結構
        export_summary = data["export_summary"]
        assert "total_elements" in export_summary
        assert "new_elements" in export_summary
        assert "updated_elements" in export_summary
        assert "unchanged_elements" in export_summary
        assert "element_breakdown" in export_summary
        
        # 檢查 export_settings 結構
        export_settings = data["export_settings"]
        assert "include_geometry" in export_settings
        assert "incremental_update" in export_settings
        assert "element_types" in export_settings
        assert "layer_filter" in export_settings
        
        # 檢查 sync_info 結構
        sync_info = data["sync_info"]
        assert "odoo_sync" in sync_info
        assert "sync_status" in sync_info
        
        # 檢查 performance 結構
        performance = data["performance"]
        assert "scan_time" in performance
        assert "export_time" in performance
        assert "total_time" in performance
        
        # 檢查時間戳格式
        assert isinstance(data["exported_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_export_to_database_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["exported_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_export_to_database_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "export_to_database called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_export_to_database_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.side_effect = Exception("測試錯誤")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in export_to_database" in call_args
        
    def test_export_to_database_drawing_name_strip(self):
        """圖面名稱去除空白測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="  測試圖面  ")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "測試圖面"
        
    def test_export_to_database_mixed_parameters(self):
        """混合參數測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [
                {"id": "AcDbCircle:1", "type": "circle", "layer": "CIRCLES", "color": 2}
            ],
            "summary": {"total_count": 1, "element_counts": {"circle": 1}}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(
            drawing_name="複合測試圖面",
            include_geometry=True,
            incremental_update=False,
            sync_to_odoo=False,
            element_types=["circle"],
            layer_filter="CIRCLES"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "複合測試圖面"
        assert result["data"]["export_settings"]["include_geometry"] == True
        assert result["data"]["export_settings"]["incremental_update"] == False
        assert result["data"]["export_settings"]["element_types"] == ["circle"]
        assert result["data"]["export_settings"]["layer_filter"] == "CIRCLES"
        assert result["data"]["sync_info"]["odoo_sync"] == False
        
    def test_export_to_database_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [{}, {}],  # 2個元素
            "summary": {"total_count": 2, "element_counts": {"line": 1, "circle": 1}}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "成功匯出 2 個元素到資料庫 (新增: 0, 更新: 0)"
        
    def test_export_to_database_default_values(self):
        """預設值測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.scan_elements.return_value = {
            "elements": [],
            "summary": {"total_count": 0}
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.export_to_database(drawing_name="測試圖面")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["export_settings"]["include_geometry"] == True
        assert result["data"]["export_settings"]["incremental_update"] == True
        assert result["data"]["export_settings"]["element_types"] == ["all"]
        assert result["data"]["export_settings"]["layer_filter"] is None
        assert result["data"]["sync_info"]["odoo_sync"] == False
        assert result["data"]["sync_info"]["sync_status"] == "not_requested"
        
    def test_export_to_database_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'export_to_database')
        assert callable(mcp_server_fastmcp.export_to_database)