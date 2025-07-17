#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
TDD Test Cases for MCP SDK Server
"""

import pytest
import asyncio
import json
from unittest.mock import Mock, patch, AsyncMock
from mcp.server.fastmcp import FastMCP
from mcp.server.stdio import stdio_server
from mcp.types import Tool, Resource
import sys
import os

# Add the project root to sys.path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class TestMCPSDKServer:
    """Test cases for MCP SDK Server implementation"""
    
    def test_server_initialization(self):
        """Test that the server can be initialized with FastMCP"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        assert server is not None
        assert server.name == "AutoCAD-Odoo Integration"
        assert server.version == "5.0"
    
    def test_server_tools_registration(self):
        """Test that tools are properly registered"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Should have basic test tools
        expected_tools = [
            'test_connection',
            'get_server_info',
            'check_autocad_status',
            'check_odoo_status'
        ]
        
        for tool_name in expected_tools:
            assert tool_name in server.get_tool_names()
    
    def test_test_connection_tool(self):
        """Test the test_connection tool"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        result = server.test_connection()
        
        assert isinstance(result, str)
        assert "MCP server is working" in result
        assert "OK" in result
    
    def test_get_server_info_tool(self):
        """Test the get_server_info tool"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        result = server.get_server_info()
        
        assert isinstance(result, dict)
        assert "name" in result
        assert "version" in result
        assert "tools_count" in result
        assert result["name"] == "AutoCAD-Odoo Integration"
        assert result["version"] == "5.0"
    
    def test_check_autocad_status_tool(self):
        """Test the check_autocad_status tool"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        result = server.check_autocad_status()
        
        assert isinstance(result, str)
        assert "AutoCAD" in result
        assert "status" in result.lower()
    
    def test_check_odoo_status_tool(self):
        """Test the check_odoo_status tool"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        result = server.check_odoo_status()
        
        assert isinstance(result, str)
        assert "Odoo" in result
        assert "status" in result.lower()
    
    def test_server_can_run_stdio(self):
        """Test that the server can run with stdio transport"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Should be able to get the FastMCP instance
        mcp_instance = server.get_mcp_instance()
        assert mcp_instance is not None
        assert hasattr(mcp_instance, 'run')
    
    def test_server_tool_schemas(self):
        """Test that tools have proper schemas"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Each tool should have a proper schema
        for tool_name in server.get_tool_names():
            tool_info = server.get_tool_info(tool_name)
            assert tool_info is not None
            assert "name" in tool_info
            assert "description" in tool_info
            # FastMCP should automatically generate schemas from function signatures
    
    @pytest.mark.asyncio
    async def test_server_async_execution(self):
        """Test that the server can handle async operations"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Test async tool execution
        result = await server.async_test_connection()
        assert isinstance(result, str)
        assert "MCP server is working" in result
    
    def test_server_error_handling(self):
        """Test that the server handles errors gracefully"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Test with invalid tool name
        with pytest.raises(ValueError):
            server.execute_tool("nonexistent_tool")
    
    def test_server_logging(self):
        """Test that the server has proper logging"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        with patch('mcp_server_sdk.logging') as mock_logging:
            mock_logger = Mock()
            mock_logging.getLogger.return_value = mock_logger
            mock_logging.basicConfig.return_value = None
            mock_logging.INFO = 20
            
            server = MCPSDKServer()
            server.test_connection()
            
            # Should have logged something
            assert mock_logger.info.called
    
    def test_server_configuration(self):
        """Test that the server can be configured"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        # Test custom configuration
        server = MCPSDKServer(
            name="Custom Server",
            version="1.0.0",
            log_level="DEBUG"
        )
        
        assert server.name == "Custom Server"
        assert server.version == "1.0.0"


class TestMCPSDKServerIntegration:
    """Integration tests for MCP SDK Server"""
    
    @pytest.mark.asyncio
    async def test_full_server_lifecycle(self):
        """Test the full server lifecycle"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Should be able to start and stop the server
        await server.start()
        assert server.is_running()
        
        await server.stop()
        assert not server.is_running()
    
    @pytest.mark.asyncio
    async def test_server_with_mock_stdio(self):
        """Test server with mocked stdio"""
        # RED: Test should fail before implementation
        from mcp_server_sdk import MCPSDKServer
        
        server = MCPSDKServer()
        
        # Mock stdio streams
        with patch('sys.stdin') as mock_stdin, \
             patch('sys.stdout') as mock_stdout:
            
            # Test server with mocked streams
            mock_stdin.readline.return_value = '{"jsonrpc": "2.0", "method": "initialize", "id": 1}\n'
            
            # Should handle the request
            result = await server.handle_stdio_request()
            assert result is not None


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v"])