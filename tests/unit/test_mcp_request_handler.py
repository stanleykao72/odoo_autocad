# -*- coding: utf-8 -*-
"""
TDD測試：MCP請求處理器

測試MCP工具的註冊、路由和執行機制
遵循CLAUDE.md的TDD方法論：Red → Green → Refactor
"""
import pytest
from unittest.mock import Mock, patch, AsyncMock
import json
import asyncio


class TestMCPRequestHandler:
    """測試MCP請求處理器的核心功能"""
    
    def test_mcp_request_handler_class_exists(self):
        """測試MCPRequestHandler類別是否存在"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        assert MCPRequestHandler is not None
    
    def test_mcp_request_handler_initialization(self):
        """測試MCPRequestHandler初始化"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        
        # Mock dependencies
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_log_util = Mock()
        
        handler = MCPRequestHandler(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util, 
            log_util=mock_log_util
        )
        
        assert handler.autocad_util == mock_autocad_util
        assert handler.odoo_util == mock_odoo_util
        assert handler.log_util == mock_log_util
    
    def test_register_tool_method_exists(self):
        """測試register_tool方法是否存在"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        assert hasattr(MCPRequestHandler, 'register_tool')
    
    def test_process_request_method_exists(self):
        """測試process_request方法是否存在"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        assert hasattr(MCPRequestHandler, 'process_request')
    
    def test_tool_registration(self):
        """測試工具註冊機制"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        
        handler = MCPRequestHandler(Mock(), Mock(), Mock())
        
        # 測試工具註冊
        def dummy_tool():
            return "test_result"
        
        handler.register_tool("test_tool", dummy_tool, "Test tool description")
        
        # 驗證工具已註冊
        assert "test_tool" in handler.tools
        assert handler.tools["test_tool"]["function"] == dummy_tool
        assert handler.tools["test_tool"]["description"] == "Test tool description"


class TestAutoCADMCPTools:
    """測試AutoCAD相關的MCP工具"""
    
    @pytest.fixture
    def mock_dependencies(self):
        """建立模擬的依賴"""
        return {
            'autocad_util': Mock(),
            'odoo_util': Mock(),
            'log_util': Mock()
        }
    
    @pytest.fixture
    def handler(self, mock_dependencies):
        """建立處理器實例"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        return MCPRequestHandler(**mock_dependencies)
    
    def test_scan_all_entities_tool_exists(self, handler):
        """測試scan_all_entities工具是否存在"""
        assert "scan_all_entities" in handler.tools
    
    def test_get_table_data_tool_exists(self, handler):
        """測試get_table_data工具是否存在"""
        assert "get_table_data" in handler.tools
    
    def test_scan_all_entities_execution(self, handler, mock_dependencies):
        """測試scan_all_entities工具執行"""
        # 設置模擬回傳值
        mock_dependencies['autocad_util'].scan_entities.return_value = [
            {'type': 'LINE', 'layer': 'Layer1'},
            {'type': 'CIRCLE', 'layer': 'Layer2'}
        ]
        
        # 執行工具
        result = handler.execute_tool("scan_all_entities", {})
        
        # 驗證結果
        assert result is not None
        assert isinstance(result, dict)
        assert 'entities' in result
    
    def test_get_table_data_execution(self, handler, mock_dependencies):
        """測試get_table_data工具執行"""
        # 設置模擬回傳值
        mock_dependencies['autocad_util'].get_layouts_values.return_value = {
            'tables': [
                {'name': 'BOQ_Table', 'data': [['Item', 'Qty'], ['Pipe', '100']]}
            ]
        }
        
        # 執行工具
        result = handler.execute_tool("get_table_data", {'layout': 'Model'})
        
        # 驗證結果
        assert result is not None
        assert 'tables' in result
        mock_dependencies['autocad_util'].get_layouts_values.assert_called_once()


class TestOdooMCPTools:
    """測試Odoo相關的MCP工具"""
    
    @pytest.fixture
    def handler_with_odoo_tools(self):
        """建立包含Odoo工具的處理器"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        return MCPRequestHandler(Mock(), Mock(), Mock())
    
    def test_get_products_by_query_tool_exists(self, handler_with_odoo_tools):
        """測試get_products_by_query工具是否存在"""
        assert "get_products_by_query" in handler_with_odoo_tools.tools
    
    def test_push_boq_to_project_tool_exists(self, handler_with_odoo_tools):
        """測試push_boq_to_project工具是否存在"""
        assert "push_boq_to_project" in handler_with_odoo_tools.tools
    
    def test_get_products_execution(self, handler_with_odoo_tools):
        """測試產品查詢工具執行"""
        # 設置模擬Odoo連接
        handler_with_odoo_tools.odoo_util.search_products.return_value = [
            {'id': 1, 'name': 'Pipe 100mm', 'code': 'P100'},
            {'id': 2, 'name': 'Valve 50mm', 'code': 'V050'}
        ]
        
        # 執行工具
        result = handler_with_odoo_tools.execute_tool("get_products_by_query", {
            'query': 'pipe'
        })
        
        # 驗證結果
        assert result is not None
        assert 'products' in result
        assert len(result['products']) == 2


class TestMCPProtocolHandling:
    """測試MCP協議處理"""
    
    @pytest.fixture
    def json_rpc_request(self):
        """建立標準的JSON-RPC MCP請求"""
        return {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": "scan_all_entities",
                "arguments": {}
            }
        }
    
    def test_process_json_rpc_request(self):
        """測試JSON-RPC請求處理"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        
        handler = MCPRequestHandler(Mock(), Mock(), Mock())
        
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": "scan_all_entities",
                "arguments": {}
            }
        }
        
        # 執行請求處理
        response = handler.process_request(request)
        
        # 驗證回應格式
        assert response is not None
        assert "jsonrpc" in response
        assert response["jsonrpc"] == "2.0"
        assert "id" in response
        assert response["id"] == 1
    
    def test_tool_list_request(self):
        """測試工具列表請求"""
        from ai_assistant.mcp_request_handler import MCPRequestHandler
        
        handler = MCPRequestHandler(Mock(), Mock(), Mock())
        
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/list",
            "params": {}
        }
        
        response = handler.process_request(request)
        
        # 驗證工具列表回應
        assert response is not None
        assert "result" in response
        assert "tools" in response["result"]
        assert isinstance(response["result"]["tools"], list)