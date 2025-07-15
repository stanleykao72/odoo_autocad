# -*- coding: utf-8 -*-
"""
Unit tests for MCP Server Manager
Following TDD principles: Red → Green → Refactor

These tests define the expected behavior of the MCP server manager
before implementation (Red phase).
"""
import pytest
import threading
import time
import socket
import json
from unittest.mock import Mock, patch, MagicMock

# This import will fail initially - that's the Red phase of TDD!
# We write the test first, then implement the functionality
from ai_assistant.mcp_server_manager import MCPServerManager


class TestMCPServerManagerInitialization:
    """Test MCP Server Manager initialization and basic properties"""
    
    def test_mcp_server_manager_can_be_created(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that MCPServerManager can be instantiated with required dependencies"""
        # Red: This test will fail because MCPServerManager doesn't exist yet
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        assert manager is not None
        assert manager.autocad_util == mock_autocad_util
        assert manager.odoo_util == mock_odoo_util
        assert manager.log_util == mock_log_util
    
    def test_mcp_server_manager_has_default_configuration(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that MCPServerManager has sensible default configuration"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util, 
            log_util=mock_log_util
        )
        
        # Expected default values from MCP_INTEGRATION_PLAN.md
        assert manager.tcp_port == 8000
        assert manager.pipe_name == r'\\.\pipe\odoo_autocad_mcp'
        assert manager.is_tcp_running == False
        assert manager.is_pipe_running == False
    
    def test_mcp_server_manager_accepts_custom_configuration(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that MCPServerManager accepts custom TCP port and pipe name"""
        custom_port = 9000
        custom_pipe = r'\\.\pipe\custom_mcp'
        
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util,
            tcp_port=custom_port,
            pipe_name=custom_pipe
        )
        
        assert manager.tcp_port == custom_port
        assert manager.pipe_name == custom_pipe


class TestMCPServerManagerTCPServer:
    """Test TCP Socket server functionality"""
    
    def test_can_start_tcp_server(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that TCP server can be started"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util,
            tcp_port=8001  # Use different port to avoid conflicts
        )
        
        # Mock socket to avoid actual network operations in tests
        with patch('socket.socket') as mock_socket:
            mock_socket_instance = Mock()
            mock_socket.return_value = mock_socket_instance
            
            manager.start_tcp_server()
            
            # Verify socket operations were called correctly
            mock_socket_instance.bind.assert_called_once_with(('localhost', 8001))
            mock_socket_instance.listen.assert_called_once_with(5)
            assert manager.is_tcp_running == True
    
    def test_can_stop_tcp_server(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that TCP server can be stopped"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        with patch('socket.socket'):
            manager.start_tcp_server()
            assert manager.is_tcp_running == True
            
            manager.stop_all_servers()
            assert manager.is_tcp_running == False
    
    def test_tcp_server_handles_connection_failure(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that TCP server handles connection failures gracefully"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        # Mock socket to raise an exception
        with patch('socket.socket') as mock_socket:
            mock_socket_instance = Mock()
            mock_socket.return_value = mock_socket_instance
            mock_socket_instance.bind.side_effect = OSError("Port already in use")
            
            with pytest.raises(OSError):
                manager.start_tcp_server()
            
            assert manager.is_tcp_running == False


class TestMCPServerManagerNamedPipeServer:
    """Test Named Pipe server functionality"""
    
    @pytest.mark.skipif(not hasattr(__import__('sys'), 'platform') or 'win' not in __import__('sys').platform,
                       reason="Named pipes are Windows-specific")
    def test_can_start_pipe_server(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that Named Pipe server can be started"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        # Mock win32pipe operations
        with patch('win32pipe.CreateNamedPipe') as mock_create_pipe:
            mock_create_pipe.return_value = 123  # Mock pipe handle
            
            manager.start_pipe_server()
            
            assert manager.is_pipe_running == True
    
    def test_can_stop_pipe_server(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that Named Pipe server can be stopped"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        with patch('win32pipe.CreateNamedPipe'):
            manager.start_pipe_server()
            assert manager.is_pipe_running == True
            
            manager.stop_all_servers()
            assert manager.is_pipe_running == False


class TestMCPServerManagerDualCommunication:
    """Test dual communication (TCP + Named Pipe) functionality"""
    
    def test_can_start_all_servers(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that both TCP and Pipe servers can be started simultaneously"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        with patch('socket.socket'), patch('win32pipe.CreateNamedPipe'):
            manager.start_all_servers()
            
            assert manager.is_tcp_running == True
            assert manager.is_pipe_running == True
            assert manager.is_running() == True
    
    def test_can_stop_all_servers(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that both servers can be stopped"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        with patch('socket.socket'), patch('win32pipe.CreateNamedPipe'):
            manager.start_all_servers()
            manager.stop_all_servers()
            
            assert manager.is_tcp_running == False
            assert manager.is_pipe_running == False
            assert manager.is_running() == False
    
    def test_is_running_returns_true_if_any_server_running(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that is_running() returns True if any server is running"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util
        )
        
        # Test with only TCP running
        manager.is_tcp_running = True
        manager.is_pipe_running = False
        assert manager.is_running() == True
        
        # Test with only Pipe running
        manager.is_tcp_running = False
        manager.is_pipe_running = True
        assert manager.is_running() == True
        
        # Test with neither running
        manager.is_tcp_running = False
        manager.is_pipe_running = False
        assert manager.is_running() == False


class TestMCPServerManagerUtilityMethods:
    """Test utility methods for server information"""
    
    def test_get_tcp_port_returns_port_when_running(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that get_tcp_port returns port number when TCP server is running"""
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util,
            tcp_port=8002
        )
        
        manager.is_tcp_running = True
        assert manager.get_tcp_port() == 8002
        
        manager.is_tcp_running = False
        assert manager.get_tcp_port() is None
    
    def test_get_pipe_name_returns_name_when_running(self, mock_autocad_util, mock_odoo_util, mock_log_util):
        """Test that get_pipe_name returns pipe name when Pipe server is running"""
        pipe_name = r'\\.\pipe\test_mcp'
        manager = MCPServerManager(
            autocad_util=mock_autocad_util,
            odoo_util=mock_odoo_util,
            log_util=mock_log_util,
            pipe_name=pipe_name
        )
        
        manager.is_pipe_running = True
        assert manager.get_pipe_name() == pipe_name
        
        manager.is_pipe_running = False
        assert manager.get_pipe_name() is None