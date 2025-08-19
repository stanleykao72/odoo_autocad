# -*- coding: utf-8 -*-
"""
Test for switch_to_layout function fix
"""

import pytest
import sys
import os
from unittest.mock import Mock, patch, MagicMock

# Add project root to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager


class TestSwitchToLayoutFix:
    """Test the improved switch_to_layout function"""
    
    def setup_method(self):
        """Setup mock environment"""
        self.mock_autocad_util = Mock()
        self.mock_odoo_util = Mock()
        
        # Setup AutoCAD connection
        self.mock_autocad_util.connected_autocad.return_value = True
        
        # Create mock layouts
        self.mock_layouts = []
        layout_names = ["Model", "Layout1", "S405-201", "E171-219"]
        
        for i, name in enumerate(layout_names):
            mock_layout = Mock()
            mock_layout.Name = name
            mock_layout.TabOrder = i
            self.mock_layouts.append(mock_layout)
        
        # Setup document with layouts
        self.mock_doc = Mock()
        self.mock_doc.Layouts = self.mock_layouts
        self.mock_autocad_util.doc = self.mock_doc
        
        # Setup active layout getter
        self.mock_autocad_util.get_active_layout.return_value = self.mock_layouts[0]  # Default to Model
        
        # Create manager
        self.manager = MCPSSEManager(port=8085, autocad_util=self.mock_autocad_util, odoo_util=self.mock_odoo_util)
    
    def test_switch_to_layout_exact_match(self):
        """Test switching to layout with exact name match"""
        # Setup: Target layout exists
        target_layout = next(l for l in self.mock_layouts if l.Name == "S405-201")
        
        # Mock the get_active_layout to return the target after switch
        def mock_get_active_layout():
            return target_layout
        
        self.mock_autocad_util.get_active_layout.side_effect = [
            self.mock_layouts[0],  # Initial call (current layout)
            target_layout          # After switch call
        ]
        
        # Test the function through the server tools
        # We need to access the registered tool function
        switch_tool = None
        for tool in self.manager.server._tools:
            if tool['name'] == 'switch_to_layout':
                switch_tool = tool['function']
                break
        
        assert switch_tool is not None, "switch_to_layout tool not found"
        
        # Execute
        result = switch_tool("S405-201")
        
        # Assert
        assert result["success"] == True
        assert result["message"] == "成功切換到 layout: S405-201"
        assert result["current_layout"]["name"] == "S405-201"
    
    def test_switch_to_layout_case_insensitive(self):
        """Test switching to layout with case insensitive match"""
        # Setup: Target layout exists but with different case
        target_layout = next(l for l in self.mock_layouts if l.Name == "Layout1")
        
        self.mock_autocad_util.get_active_layout.side_effect = [
            self.mock_layouts[0],  # Initial call
            target_layout          # After switch call
        ]
        
        # Get the tool function
        switch_tool = None
        for tool in self.manager.server._tools:
            if tool['name'] == 'switch_to_layout':
                switch_tool = tool['function']
                break
        
        # Execute with different case
        result = switch_tool("layout1")
        
        # Assert
        assert result["success"] == True
        assert result["current_layout"]["name"] == "Layout1"
    
    def test_switch_to_layout_not_found(self):
        """Test switching to non-existent layout"""
        # Get the tool function
        switch_tool = None
        for tool in self.manager.server._tools:
            if tool['name'] == 'switch_to_layout':
                switch_tool = tool['function']
                break
        
        # Execute with non-existent layout
        result = switch_tool("NonExistentLayout")
        
        # Assert
        assert result["success"] == False
        assert "找不到名為 'NonExistentLayout' 的 layout" in result["error"]
        assert "available_layouts" in result
        assert len(result["available_layouts"]) == 4
        assert "S405-201" in result["available_layouts"]
    
    def test_switch_to_layout_autocad_not_connected(self):
        """Test switching when AutoCAD is not connected"""
        # Setup: AutoCAD not connected
        self.mock_autocad_util.connected_autocad.return_value = False
        
        # Get the tool function
        switch_tool = None
        for tool in self.manager.server._tools:
            if tool['name'] == 'switch_to_layout':
                switch_tool = tool['function']
                break
        
        # Execute
        result = switch_tool("S405-201")
        
        # Assert
        assert result["success"] == False
        assert result["error"] == "AutoCAD 未連接"
    
    def test_switch_to_layout_with_debug_info(self):
        """Test that the improved function provides useful debug information"""
        # Setup: Layout exists but switching fails
        target_layout = next(l for l in self.mock_layouts if l.Name == "S405-201")
        
        # Mock switching to fail - ActiveLayout assignment doesn't work
        def mock_setattr(attr, value):
            if attr == "ActiveLayout":
                raise Exception("Permission denied")
        
        self.mock_doc.__setattr__ = mock_setattr
        
        # Mock get_active_layout to return original layout (indicating switch failed)
        self.mock_autocad_util.get_active_layout.side_effect = [
            self.mock_layouts[0],  # Initial call
            self.mock_layouts[0]   # After switch - still the same (failed)
        ]
        
        # Get the tool function
        switch_tool = None
        for tool in self.manager.server._tools:
            if tool['name'] == 'switch_to_layout':
                switch_tool = tool['function']
                break
        
        # Execute
        result = switch_tool("S405-201")
        
        # Assert
        assert result["success"] == False
        assert "debug_info" in result
        assert "target_layout_found" in result["debug_info"]
        assert result["debug_info"]["target_layout_found"] == "S405-201"
    
    def test_switch_to_layout_multiple_methods(self):
        """Test that multiple switching methods are attempted"""
        # Setup: First method fails, second succeeds
        target_layout = next(l for l in self.mock_layouts if l.Name == "S405-201")
        
        # Mock the document and app
        self.mock_app = Mock()
        self.mock_app.ActiveDocument = self.mock_doc
        self.mock_autocad_util.app = self.mock_app
        
        # Make method 1 fail, method 2 succeed
        def mock_doc_setattr(attr, value):
            if attr == "ActiveLayout":
                raise Exception("Method 1 failed")
        
        self.mock_doc.__setattr__ = mock_doc_setattr
        
        # Method 2 should work
        self.mock_doc.ActiveLayout = target_layout  # This will be set by method 2
        
        # Mock successful switch verification
        self.mock_autocad_util.get_active_layout.side_effect = [
            self.mock_layouts[0],  # Initial call
            target_layout          # After switch call
        ]
        
        # Get the tool function
        switch_tool = None
        for tool in self.manager.server._tools:
            if tool['name'] == 'switch_to_layout':
                switch_tool = tool['function']
                break
        
        # Execute
        result = switch_tool("S405-201")
        
        # Should succeed via method 2
        assert result["success"] == True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])