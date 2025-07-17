# -*- coding: utf-8 -*-
"""
TDD測試文件：generate_boq_from_drawing 功能
生成工程量清單 (BOQ) 的完整測試套件

遵循 TDD 原則：Red -> Green -> Refactor
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime
import json

import mcp_server_fastmcp


class TestGenerateBoqFromDrawing:
    """generate_boq_from_drawing 功能的 TDD 測試類"""
    
    def test_generate_boq_from_drawing_function_exists(self):
        """驗證函數存在"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'generate_boq_from_drawing')
        assert callable(mcp_server_fastmcp.generate_boq_from_drawing)
        
    def test_generate_boq_from_drawing_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = {
            "id": 456,
            "name": "混凝土 C25",
            "unit_price": 3200.0,
            "unit": "m³"
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {
                            "id": "AcDbSolid:123",
                            "type": "solid",
                            "layer": "CONCRETE",
                            "volume": 25.5,
                            "properties": {"material": "concrete"}
                        }
                    ],
                    "summary": {"total_count": 1}
                }
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                project_id=123
            )
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["drawing_name"] == "測試圖面"
        assert result["data"]["boq_info"]["project_id"] == 123
        assert result["data"]["boq_summary"]["total_items"] > 0
        assert result["data"]["boq_summary"]["total_amount"] > 0
        assert result["data"]["boq_summary"]["currency"] == "TWD"
        assert "boq_items" in result["data"]
        assert "calculation_details" in result["data"]
        assert "odoo_integration" in result["data"]
        assert "export_info" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_generate_boq_from_drawing_parameter_validation_empty_drawing_name(self):
        """參數驗證測試 - 空白圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(drawing_name="")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        assert "圖面名稱不能為空" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_whitespace_drawing_name(self):
        """參數驗證測試 - 只有空白的圖面名稱"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(drawing_name="   ")
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        assert "圖面名稱不能為空" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_invalid_project_id_negative(self):
        """參數驗證測試 - 負數專案ID"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面",
            project_id=-1
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_ID"
        assert "專案ID必須是正整數" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_invalid_project_id_zero(self):
        """參數驗證測試 - 零專案ID"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面",
            project_id=0
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_ID"
        assert "專案ID必須是正整數" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_invalid_project_id_type(self):
        """參數驗證測試 - 非整數專案ID"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面",
            project_id="invalid"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_PROJECT_ID_TYPE"
        assert "專案ID必須是整數" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_invalid_output_format(self):
        """參數驗證測試 - 無效的輸出格式"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面",
            output_format="invalid"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_OUTPUT_FORMAT"
        assert "無效的輸出格式" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_valid_output_formats(self):
        """參數驗證測試 - 有效的輸出格式"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        valid_formats = ["odoo", "excel", "json"]
        
        for format_type in valid_formats:
            with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
                mock_scan.return_value = {
                    "status": "success",
                    "data": {"elements": [], "summary": {"total_count": 0}}
                }
                
                # Act
                result = mcp_server_fastmcp.generate_boq_from_drawing(
                    drawing_name="測試圖面",
                    output_format=format_type
                )
                
                # Assert
                assert result["status"] == "success"
                assert result["data"]["export_info"]["format"] == format_type
                
    def test_generate_boq_from_drawing_parameter_validation_invalid_currency(self):
        """參數驗證測試 - 無效的貨幣"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面",
            currency="INVALID"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_CURRENCY"
        assert "無效的貨幣代碼" in result["message"]
        
    def test_generate_boq_from_drawing_parameter_validation_valid_currencies(self):
        """參數驗證測試 - 有效的貨幣代碼"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        valid_currencies = ["TWD", "USD", "EUR", "JPY", "CNY"]
        
        for currency in valid_currencies:
            with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
                mock_scan.return_value = {
                    "status": "success",
                    "data": {"elements": [], "summary": {"total_count": 0}}
                }
                
                # Act
                result = mcp_server_fastmcp.generate_boq_from_drawing(
                    drawing_name="測試圖面",
                    currency=currency
                )
                
                # Assert
                assert result["status"] == "success"
                assert result["data"]["boq_summary"]["currency"] == currency
                
    def test_generate_boq_from_drawing_autocad_connection_error(self):
        """AutoCAD連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        mcp_server_fastmcp._odoo_util = Mock()
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "AutoCAD 連接未建立" in result["message"]
        
    def test_generate_boq_from_drawing_odoo_connection_error(self):
        """Odoo連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = Mock()
        mcp_server_fastmcp._odoo_util = None
        
        # Act
        result = mcp_server_fastmcp.generate_boq_from_drawing(
            drawing_name="測試圖面"
        )
        
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "ODOO_NOT_CONNECTED"
        assert "Odoo 連接未建立" in result["message"]
        
    def test_generate_boq_from_drawing_scan_elements_failed(self):
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
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面"
            )
            
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "SCAN_ELEMENTS_FAILED"
        assert "掃描失敗" in result["message"]
        
    def test_generate_boq_from_drawing_with_element_types_filter(self):
        """元素類型過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {"id": "1", "type": "line", "layer": "WALL"},
                        {"id": "2", "type": "circle", "layer": "COLUMN"},
                        {"id": "3", "type": "solid", "layer": "CONCRETE"}
                    ],
                    "summary": {"total_count": 3}
                }
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                element_types=["line", "solid"]
            )
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["calculation_details"]["total_elements_processed"] == 2
        
    def test_generate_boq_from_drawing_with_layer_filter(self):
        """圖層過濾測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {"id": "1", "type": "line", "layer": "WALL"},
                        {"id": "2", "type": "circle", "layer": "COLUMN"},
                        {"id": "3", "type": "solid", "layer": "CONCRETE"}
                    ],
                    "summary": {"total_count": 3}
                }
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                layer_filter="CONCRETE"
            )
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["calculation_details"]["total_elements_processed"] == 1
        
    def test_generate_boq_from_drawing_with_custom_calculation_rules(self):
        """自訂計算規則測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        custom_rules = {
            "concrete_density": 2400,
            "rebar_weight_per_meter": 0.617,
            "formwork_wastage": 0.1
        }
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                calculation_rules=custom_rules
            )
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["calculation_rules"] == custom_rules
        
    def test_generate_boq_from_drawing_without_autocad_data(self):
        """不包含AutoCAD資料測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                include_autocad_data=False
            )
            
        # Assert
        assert result["status"] == "success"
        for item in result["data"]["boq_items"]:
            assert "autocad_elements" not in item or len(item["autocad_elements"]) == 0
            
    def test_generate_boq_from_drawing_with_drawing_name_strip(self):
        """圖面名稱去除空格測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="  測試圖面  "
            )
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["drawing_name"] == "測試圖面"
        
    def test_generate_boq_from_drawing_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.side_effect = Exception("產品查詢錯誤")
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [{"id": "test", "type": "line", "layer": "TEST"}], "summary": {"total_count": 1}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面"
            )
            
        # Assert
        assert result["status"] == "error"
        assert result["error_code"] == "BOQ_GENERATION_ERROR"
        assert "產品查詢錯誤" in result["message"]
        
    def test_generate_boq_from_drawing_return_format_complete(self):
        """返回格式完整性測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = {
            "id": 456,
            "name": "混凝土 C25",
            "unit_price": 3200.0,
            "unit": "m³"
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {
                            "id": "AcDbSolid:123",
                            "type": "solid",
                            "layer": "CONCRETE",
                            "volume": 25.5,
                            "properties": {"material": "concrete"}
                        }
                    ],
                    "summary": {"total_count": 1}
                }
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面",
                project_id=123
            )
            
        # Assert
        assert result["status"] == "success"
        assert "data" in result
        assert "boq_info" in result["data"]
        assert "boq_summary" in result["data"]
        assert "boq_items" in result["data"]
        assert "calculation_details" in result["data"]
        assert "odoo_integration" in result["data"]
        assert "export_info" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
        # 檢查 boq_info 結構
        boq_info = result["data"]["boq_info"]
        assert "drawing_name" in boq_info
        assert "project_id" in boq_info
        assert "generated_at" in boq_info
        assert "calculation_rules" in boq_info
        assert "currency" in boq_info
        
        # 檢查 boq_summary 結構
        boq_summary = result["data"]["boq_summary"]
        assert "total_items" in boq_summary
        assert "total_quantity" in boq_summary
        assert "total_amount" in boq_summary
        assert "currency" in boq_summary
        assert "item_categories" in boq_summary
        
    def test_generate_boq_from_drawing_timestamp_format(self):
        """時間戳格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            with patch('mcp_server_fastmcp.datetime') as mock_datetime:
                mock_now = Mock()
                mock_now.isoformat.return_value = "2025-07-16T10:30:00.123456"
                mock_datetime.now.return_value = mock_now
                
                # Act
                result = mcp_server_fastmcp.generate_boq_from_drawing(
                    drawing_name="測試圖面"
                )
                
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["generated_at"] == "2025-07-16T10:30:00.123456"
        assert result["timestamp"] == "2025-07-16T10:30:00.123456"
        
    @patch('mcp_server_fastmcp.logger')
    def test_generate_boq_from_drawing_logging_success(self, mock_logger):
        """驗證成功情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面"
            )
            
        # Assert
        assert result["status"] == "success"
        mock_logger.info.assert_called()
        
    @patch('mcp_server_fastmcp.logger')
    def test_generate_boq_from_drawing_logging_error(self, mock_logger):
        """驗證錯誤情況的日誌記錄"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.side_effect = Exception("測試錯誤")
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [{"id": "test", "type": "line", "layer": "TEST"}], "summary": {"total_count": 1}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面"
            )
            
        # Assert
        assert result["status"] == "error"
        mock_logger.error.assert_called()
        
    def test_generate_boq_from_drawing_success_message_format(self):
        """成功訊息格式測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = {
            "id": 456,
            "name": "混凝土 C25",
            "unit_price": 3200.0,
            "unit": "m³"
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {
                    "elements": [
                        {
                            "id": "AcDbSolid:123",
                            "type": "solid",
                            "layer": "CONCRETE",
                            "volume": 25.5,
                            "properties": {"material": "concrete"}
                        }
                    ],
                    "summary": {"total_count": 1}
                }
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面"
            )
            
        # Assert
        assert result["status"] == "success"
        assert "成功生成BOQ" in result["message"]
        assert "測試圖面" in result["message"]
        assert "項工程量清單" in result["message"]
        assert "總金額" in result["message"]
        
    def test_generate_boq_from_drawing_default_values(self):
        """預設值測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.get_product_by_name.return_value = None
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        with patch('mcp_server_fastmcp.scan_elements') as mock_scan:
            mock_scan.return_value = {
                "status": "success",
                "data": {"elements": [], "summary": {"total_count": 0}}
            }
            
            # Act
            result = mcp_server_fastmcp.generate_boq_from_drawing(
                drawing_name="測試圖面"
            )
            
        # Assert
        assert result["status"] == "success"
        assert result["data"]["boq_info"]["project_id"] is None
        assert result["data"]["boq_summary"]["currency"] == "TWD"
        assert result["data"]["export_info"]["format"] == "odoo"
        assert result["data"]["boq_info"]["calculation_rules"] == "standard"