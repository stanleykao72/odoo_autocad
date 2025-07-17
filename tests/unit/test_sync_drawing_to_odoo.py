"""
TDD 測試：sync_drawing_to_odoo MCP 工具

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


class TestSyncDrawingToOdoo:
    """sync_drawing_to_odoo 功能的 TDD 測試"""
    
    def setup_method(self):
        """每個測試方法前的設定"""
        # 重設全局變數
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = None
    
    def test_sync_drawing_to_odoo_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 23,
            "failed_count": 2,
            "element_breakdown": {"line": 10, "circle": 5, "text": 6, "dimension": 2}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 8,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Mock scan_elements function
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {"id": "elem1", "type": "line", "layer": "0"},
                        {"id": "elem2", "type": "circle", "layer": "0"}
                    ],
                    "summary": {"total_count": 2}
                }
            }
            
            # Mock generate_boq_from_drawing function
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {
                        "boq_summary": {
                            "total_items": 15,
                            "total_amount": 125000.0,
                            "currency": "TWD"
                        }
                    }
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面",
                    project_name="測試專案"
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["project_info"]["project_id"] == 123
        assert result["data"]["project_info"]["project_name"] == "測試專案"
        assert result["data"]["project_info"]["created_new"] == True
        assert result["data"]["sync_summary"]["synced_elements"] == 23
        assert result["data"]["sync_summary"]["failed_elements"] == 2
        assert result["data"]["sync_summary"]["sync_mode"] == "full"
        assert "sync_details" in result["data"]
        assert "odoo_urls" in result["data"]
        assert "sync_log" in result["data"]
        assert "synced_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_sync_drawing_to_odoo_with_project_id(self):
        """使用專案ID測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 456,
            "project_name": "現有專案",
            "created_new": False,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 15,
            "failed_count": 0,
            "element_breakdown": {"line": 8, "circle": 7}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 5,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Mock functions
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 10, "total_amount": 80000.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面",
                    project_id=456
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["project_info"]["project_id"] == 456
        assert result["data"]["project_info"]["created_new"] == False
        
    def test_sync_drawing_to_odoo_parameter_validation_empty_drawing_name(self):
        """參數驗證測試 - 空白圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(drawing_name="")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        assert "圖面名稱不能為空" in result["message"]
        
    def test_sync_drawing_to_odoo_parameter_validation_whitespace_drawing_name(self):
        """參數驗證測試 - 只有空白的圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(drawing_name="   ")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        
    def test_sync_drawing_to_odoo_parameter_validation_invalid_sync_mode(self):
        """參數驗證測試 - 無效同步模式"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            sync_mode="invalid_mode"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_SYNC_MODE"
        assert "無效的同步模式" in result["message"]
        
    def test_sync_drawing_to_odoo_parameter_validation_valid_sync_modes(self):
        """參數驗證測試 - 有效同步模式"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 10,
            "failed_count": 0,
            "element_breakdown": {}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 5,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        valid_modes = ["full", "incremental", "elements_only"]
        
        for sync_mode in valid_modes:
            with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
                mock_scan.return_value = {
                    "status": "success",
                    "data": {"elements": [], "summary": {"total_count": 0}}
                }
                
                with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                    mock_boq.return_value = {
                        "status": "success",
                        "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                    }
                    
                    # Act
                    result = mcp_server_fastmcp.sync_drawing_to_odoo(
                        drawing_name="測試圖面",
                        sync_mode=sync_mode
                    )
                    
                    # Assert
                    assert result["status"] == "success"
                    assert result["data"]["sync_summary"]["sync_mode"] == sync_mode
        
    def test_sync_drawing_to_odoo_parameter_validation_conflicting_project_params(self):
        """參數驗證測試 - 專案參數衝突"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_name="專案名稱",
            project_id=123
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "CONFLICTING_PROJECT_PARAMS"
        assert "不能同時指定專案名稱和專案ID" in result["message"]
        
    def test_sync_drawing_to_odoo_parameter_validation_invalid_project_name_empty(self):
        """參數驗證測試 - 專案名稱為空"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_name=""
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_NAME"
        assert "專案名稱不能為空" in result["message"]
        
    def test_sync_drawing_to_odoo_parameter_validation_invalid_project_name_whitespace(self):
        """參數驗證測試 - 專案名稱只有空白"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_name="   "
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_NAME"
        
    def test_sync_drawing_to_odoo_parameter_validation_invalid_project_id_negative(self):
        """參數驗證測試 - 專案ID為負數"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_id=-1
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_ID"
        assert "專案ID必須是正整數" in result["message"]
        
    def test_sync_drawing_to_odoo_parameter_validation_invalid_project_id_zero(self):
        """參數驗證測試 - 專案ID為零"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_id=0
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_ID"
        
    def test_sync_drawing_to_odoo_parameter_validation_invalid_project_id_type(self):
        """參數驗證測試 - 專案ID類型無效"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_id="not_a_number"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_ID_TYPE"
        assert "專案ID必須是整數" in result["message"]
        
    def test_sync_drawing_to_odoo_parameter_validation_project_id_conversion(self):
        """參數驗證測試 - 專案ID轉換"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": False,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 0,
            "failed_count": 0,
            "element_breakdown": {}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 0,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面",
                    project_id="123"  # 字串數字
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["project_info"]["project_id"] == 123
        
    def test_sync_drawing_to_odoo_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = Mock()
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_sync_drawing_to_odoo_odoo_connection_error(self):
        """Odoo 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = Mock()
        mcp_server_fastmcp._odoo_util = None
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "ODOO_NOT_CONNECTED"
        assert "Odoo 連接未建立" in result["message"]
        
    def test_sync_drawing_to_odoo_scan_elements_failed(self):
        """掃描元素失敗測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "error",
                "message": "掃描失敗"
            }
            
            # Act
            result = mcp_server_fastmcp.sync_drawing_to_odoo(
                drawing_name="測試圖面",
                sync_elements=True
            )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "SCAN_ELEMENTS_FAILED"
        assert "無法獲取圖面元素" in result["message"]
        
    def test_sync_drawing_to_odoo_elements_only_mode(self):
        """僅元素同步模式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 10,
            "failed_count": 0,
            "element_breakdown": {"line": 5, "circle": 5}
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [{"id": "1", "type": "line"}], "summary": {"total_count": 1}}
            }
            
            # Act
            result = mcp_server_fastmcp.sync_drawing_to_odoo(
                drawing_name="測試圖面",
                sync_mode="elements_only",
                sync_boq=False,
                sync_parameters=False
            )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_summary"]["sync_mode"] == "elements_only"
        
    def test_sync_drawing_to_odoo_incremental_mode(self):
        """增量同步模式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": False,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 5,
            "failed_count": 0,
            "element_breakdown": {"line": 3, "circle": 2}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 3,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 5, "total_amount": 50000.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面",
                    sync_mode="incremental"
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_summary"]["sync_mode"] == "incremental"
        
    def test_sync_drawing_to_odoo_with_drawing_name_strip(self):
        """圖面名稱去除空白測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試圖面",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 0,
            "failed_count": 0,
            "element_breakdown": {}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 0,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="  測試圖面  ",
                    project_name="  測試專案  "
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["drawing_name"] == "測試圖面"
        
    def test_sync_drawing_to_odoo_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.side_effect = Exception("同步錯誤")
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.sync_drawing_to_odoo(
                drawing_name="測試圖面"
            )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "DRAWING_SYNC_ERROR"
        assert "圖面同步失敗" in result["message"]
        
    def test_sync_drawing_to_odoo_return_format_complete(self):
        """驗證回傳格式完整性"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 20,
            "failed_count": 5,
            "element_breakdown": {"line": 10, "circle": 8, "text": 2}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 8,
            "error_count": 1
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [{"id": "1", "type": "line"}], "summary": {"total_count": 1}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 15, "total_amount": 125000.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面"
                )
        
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 data 結構
        data = result["data"]
        assert "project_info" in data
        assert "sync_summary" in data
        assert "sync_details" in data
        assert "odoo_urls" in data
        assert "sync_log" in data
        assert "synced_at" in data
        
        # 檢查 project_info 結構
        project_info = data["project_info"]
        assert "project_id" in project_info
        assert "project_name" in project_info
        assert "created_new" in project_info
        assert "updated_at" in project_info
        
        # 檢查 sync_summary 結構
        sync_summary = data["sync_summary"]
        assert "total_elements" in sync_summary
        assert "synced_elements" in sync_summary
        assert "failed_elements" in sync_summary
        assert "sync_mode" in sync_summary
        assert "sync_time" in sync_summary
        
        # 檢查 sync_details 結構
        sync_details = data["sync_details"]
        assert "elements_sync" in sync_details
        assert "boq_sync" in sync_details
        assert "parameters_sync" in sync_details
        
        # 檢查 odoo_urls 結構
        odoo_urls = data["odoo_urls"]
        assert "project_url" in odoo_urls
        assert "boq_url" in odoo_urls
        
        # 檢查時間戳格式
        assert isinstance(data["synced_at"], str)
        assert isinstance(result["timestamp"], str)
        
    @patch('mcp_server_fastmcp.datetime')
    def test_sync_drawing_to_odoo_timestamp_format(self, mock_datetime):
        """驗證時間戳格式"""
        # Arrange
        mock_now = Mock()
        mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
        mock_datetime.now.return_value = mock_now
        
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 0,
            "failed_count": 0,
            "element_breakdown": {}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 0,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面"
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["synced_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_sync_drawing_to_odoo_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 0,
            "failed_count": 0,
            "element_breakdown": {}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 0,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面"
                )
        
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args[0][0]
        assert "sync_drawing_to_odoo called with params" in call_args
        
    @patch('mcp_server_fastmcp.logger')
    def test_sync_drawing_to_odoo_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.side_effect = Exception("測試錯誤")
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.sync_drawing_to_odoo(
                drawing_name="測試圖面"
            )
        
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called_once()
        call_args = mock_logger.error.call_args[0][0]
        assert "Error in sync_drawing_to_odoo" in call_args
        
    def test_sync_drawing_to_odoo_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 15,
            "failed_count": 5,
            "element_breakdown": {"line": 10, "circle": 5}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 0,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [{"id": "1"}] * 20, "summary": {"total_count": 20}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面",
                    project_name="測試專案"
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["message"] == "圖面同步完成: 15/20 元素成功同步到專案 '測試專案'"
        
    def test_sync_drawing_to_odoo_default_values(self):
        """預設值測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試圖面",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 0,
            "failed_count": 0,
            "element_breakdown": {}
        }
        mock_odoo_util.sync_drawing_parameters.return_value = {
            "synced_count": 0,
            "error_count": 0
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.generate_boq_from_drawing') as mock_boq:
                mock_boq.return_value = {
                    "status": "success",
                    "data": {"boq_summary": {"total_items": 0, "total_amount": 0.0, "currency": "TWD"}}
                }
                
                # Act
                result = mcp_server_fastmcp.sync_drawing_to_odoo(
                    drawing_name="測試圖面"
                )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sync_summary"]["sync_mode"] == "full"
        assert result["data"]["project_info"]["project_name"] == "測試圖面"  # 使用圖面名稱作為專案名稱
        
    def test_sync_drawing_to_odoo_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'sync_drawing_to_odoo')
        assert callable(mcp_server_fastmcp.sync_drawing_to_odoo)