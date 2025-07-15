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
        """Test that create_ai_control_panel method exists"""
        from forms.form_main_modern import ModernFormMain
        
        # Check if the method exists
        assert hasattr(ModernFormMain, 'create_ai_control_panel'), \
            "create_ai_control_panel method should exist"
        
        # Check if it's callable
        assert callable(getattr(ModernFormMain, 'create_ai_control_panel')), \
            "create_ai_control_panel should be callable"
    
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
    
    def test_open_ai_chat_method_exists(self):
        """Test that open_ai_chat method exists"""
        from forms.form_main_modern import ModernFormMain
        
        # Check if the method exists  
        assert hasattr(ModernFormMain, 'open_ai_chat'), \
            "open_ai_chat method should exist"
        
        # Check if it's callable
        assert callable(getattr(ModernFormMain, 'open_ai_chat')), \
            "open_ai_chat should be callable"


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
        form.ai_chat_button = Mock()
        form.tcp_info_label = Mock()
        form.pipe_info_label = Mock()
        
        # Bind the real method to our mock
        form.update_mcp_status_display = ModernFormMain.update_mcp_status_display.__get__(form)
        
        # Call the method
        form.update_mcp_status_display()
        
        # Verify UI updates for running state
        form.mcp_status_label.configure.assert_called_with(text="🟢 AI助手運行中")
        form.mcp_toggle_button.configure.assert_called_with(text="⏹️ 停止AI助手")
        form.ai_chat_button.configure.assert_called_with(state="normal")
        form.tcp_info_label.configure.assert_called_with(text="TCP: localhost:8000")
        form.pipe_info_label.configure.assert_called_with(text="Pipe: 運行中")
    
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
        form.ai_chat_button = Mock()
        form.tcp_info_label = Mock()
        form.pipe_info_label = Mock()
        
        # Bind the real method to our mock
        form.update_mcp_status_display = ModernFormMain.update_mcp_status_display.__get__(form)
        
        # Call the method
        form.update_mcp_status_display()
        
        # Verify UI updates for stopped state
        form.mcp_status_label.configure.assert_called_with(text="🔴 AI助手離線")
        form.mcp_toggle_button.configure.assert_called_with(text="🚀 啟動AI助手")
        form.ai_chat_button.configure.assert_called_with(state="disabled")
        form.tcp_info_label.configure.assert_called_with(text="TCP: 未啟動")
        form.pipe_info_label.configure.assert_called_with(text="Pipe: 未啟動")
    
    def test_open_ai_chat_warns_when_server_not_running(self):
        """Test that open_ai_chat shows warning when MCP server is not running"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = Mock()
        form.mcp_server_manager.is_running.return_value = False
        form.log_util = Mock()
        
        # Bind the real method to our mock
        form.open_ai_chat = ModernFormMain.open_ai_chat.__get__(form)
        
        # Mock messagebox
        with patch('forms.form_main_modern.messagebox') as mock_messagebox:
            # Call the method
            form.open_ai_chat()
            
            # Verify warning was shown
            mock_messagebox.showwarning.assert_called_once()
            call_args = mock_messagebox.showwarning.call_args
            assert "AI助手未啟動" in call_args[0][0]
    
    def test_open_ai_chat_shows_development_message_when_running(self):
        """Test that open_ai_chat shows development message when server is running"""
        from forms.form_main_modern import ModernFormMain
        
        # Create a mock form instance
        form = Mock(spec=ModernFormMain)
        form.mcp_server_manager = Mock()
        form.mcp_server_manager.is_running.return_value = True
        form.log_util = Mock()
        
        # Bind the real method to our mock
        form.open_ai_chat = ModernFormMain.open_ai_chat.__get__(form)
        
        # Mock messagebox
        with patch('forms.form_main_modern.messagebox') as mock_messagebox:
            # Call the method
            form.open_ai_chat()
            
            # Verify development message was shown
            mock_messagebox.showinfo.assert_called_once()
            call_args = mock_messagebox.showinfo.call_args
            assert "功能開發中" in call_args[0][0]