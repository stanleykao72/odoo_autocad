# -*- coding: utf-8 -*-
"""
Unit tests for AI Control Panel
Following TDD principles: Red → Green → Refactor

These tests define the expected behavior of the AI control panel
before implementation (Red phase).
"""
import pytest
import customtkinter as ctk
from unittest.mock import Mock, patch, MagicMock


class TestAIControlPanelMethods:
    """Test AI Control Panel method existence and basic functionality"""
    
    def test_create_ai_control_panel_method_exists(self):
        """Test that create_ai_control_banner method exists (updated to banner design)"""
        from forms.form_main_modern import ModernFormMain
        
        # Check if the method exists
        assert hasattr(ModernFormMain, 'create_ai_control_banner'), \
            "create_ai_control_banner method should exist"
        
        # Check if it's callable
        assert callable(getattr(ModernFormMain, 'create_ai_control_banner')), \
            "create_ai_control_banner should be callable"
    
    def test_initialize_mcp_server_manager_method_exists(self):
        """Test that initialize_mcp_server_manager method exists"""
        from forms.form_main_modern import ModernFormMain
        
        # Check if the method exists
        assert hasattr(ModernFormMain, 'initialize_mcp_server_manager'), \
            "initialize_mcp_server_manager method should exist"
        
        # Check if it's callable
        assert callable(getattr(ModernFormMain, 'initialize_mcp_server_manager')), \
            "initialize_mcp_server_manager should be callable"
    
    def test_toggle_mcp_server_method_exists(self):
        """Test that toggle_mcp_server method exists"""
        from forms.form_main_modern import ModernFormMain
        
        # Check if the method exists
        assert hasattr(ModernFormMain, 'toggle_mcp_server'), \
            "toggle_mcp_server method should exist"
        
        # Check if it's callable
        assert callable(getattr(ModernFormMain, 'toggle_mcp_server')), \
            "toggle_mcp_server should be callable"
    
    def test_update_mcp_status_display_method_exists(self):
        """Test that update_mcp_status_display method exists"""
        from forms.form_main_modern import ModernFormMain
        
        # Check if the method exists
        assert hasattr(ModernFormMain, 'update_mcp_status_display'), \
            "update_mcp_status_display method should exist"
        
        # Check if it's callable
        assert callable(getattr(ModernFormMain, 'update_mcp_status_display')), \
            "update_mcp_status_display should be callable"
    


class TestAIControlPanelLogic:
    """Test AI Control Panel business logic with mocked dependencies"""
    
    def test_initialize_mcp_server_manager_creates_instance(self):
        """Test that initialize_mcp_server_manager creates MCP server manager instance"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = None
        form.autocad_util = Mock()
        form.odoo_util = Mock()
        form.log_util = Mock()
        
        # Bind the real method to our mock
        form.initialize_mcp_server_manager = ModernFormMain.initialize_mcp_server_manager.__get__(form)
        
        # Mock MCPServerManager import
        with patch('forms.form_main_modern.MCPServerManager') as mock_mcp_class:
            mock_instance = Mock()
            mock_mcp_class.return_value = mock_instance
            
            # Call the method
            form.initialize_mcp_server_manager()
            
            # Verify MCPServerManager was created with correct dependencies
            mock_mcp_class.assert_called_once_with(
                autocad_util=form.autocad_util,
                odoo_util=form.odoo_util,
                log_util=form.log_util
            )
            
            # Verify the instance was assigned
            assert form.mcp_server_manager == mock_instance
    
    def test_toggle_mcp_server_starts_when_stopped(self):
        """Test that toggle_mcp_server starts MCP server when stopped"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = Mock()
        form.mcp_server_manager.is_running.return_value = False
        form.log_util = Mock()
        
        # Mock the initialize and update methods
        form.initialize_mcp_server_manager = Mock()
        form.update_mcp_status_display = Mock()
        
        # Bind the real method to our mock
        form.toggle_mcp_server = ModernFormMain.toggle_mcp_server.__get__(form)
        
        # Call the method
        form.toggle_mcp_server()
        
        # Verify server was started
        form.mcp_server_manager.start_all_servers.assert_called_once()
        form.update_mcp_status_display.assert_called_once()
    
    def test_toggle_mcp_server_stops_when_running(self):
        """Test that toggle_mcp_server stops MCP server when running"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = Mock()
        form.mcp_server_manager.is_running.return_value = True
        form.log_util = Mock()
        
        # Mock the update method
        form.update_mcp_status_display = Mock()
        
        # Bind the real method to our mock
        form.toggle_mcp_server = ModernFormMain.toggle_mcp_server.__get__(form)
        
        # Call the method
        form.toggle_mcp_server()
        
        # Verify server was stopped
        form.mcp_server_manager.stop_all_servers.assert_called_once()
        form.update_mcp_status_display.assert_called_once()
    
    def test_update_mcp_status_display_shows_running_state(self):
        """Test that update_mcp_status_display correctly shows running state"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = Mock()
        form.mcp_server_manager.is_running.return_value = True
        form.mcp_server_manager.get_tcp_port.return_value = 8000
        form.mcp_server_manager.get_pipe_name.return_value = r'\\.\pipe\odoo_autocad_mcp'
        
        # Mock UI components
        form.mcp_status_label = Mock()
        form.mcp_toggle_button = Mock()
        form.tcp_info_label = Mock()
        form.pipe_info_label = None
        
        # Bind the real method to our mock
        form.update_mcp_status_display = ModernFormMain.update_mcp_status_display.__get__(form)
        
        # Call the method
        form.update_mcp_status_display()
        
        # Verify UI updates for running state (icon-only design)
        form.mcp_status_label.configure.assert_called_with(text="🟢")
        form.mcp_toggle_button.configure.assert_called_with(text="⏹️")
        form.tcp_info_label.configure.assert_called_with(text="AI: :8000")
    
    def test_update_mcp_status_display_shows_stopped_state(self):
        """Test that update_mcp_status_display correctly shows stopped state"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance  
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = Mock()
        form.mcp_server_manager.is_running.return_value = False
        
        # Mock UI components
        form.mcp_status_label = Mock()
        form.mcp_toggle_button = Mock()
        form.tcp_info_label = Mock()
        form.pipe_info_label = None
        
        # Bind the real method to our mock
        form.update_mcp_status_display = ModernFormMain.update_mcp_status_display.__get__(form)
        
        # Call the method
        form.update_mcp_status_display()
        
        # Verify UI updates for stopped state (icon-only design)
        form.mcp_status_label.configure.assert_called_with(text="🔴")
        form.mcp_toggle_button.configure.assert_called_with(text="🚀")
        form.tcp_info_label.configure.assert_called_with(text="")
    
