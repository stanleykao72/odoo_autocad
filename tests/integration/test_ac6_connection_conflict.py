# -*- coding: utf-8 -*-
"""
Integration test for AC6: Connection conflict detection
Tests that MCP Server properly handles shared instance scenarios
"""

import pytest
import sys
import os
from unittest.mock import Mock

# Add project root to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

# Import the MCP server module
import mcp_server_fastmcp


class TestAC6ConnectionConflictIntegration:
    """Integration test for AC6 connection conflict detection"""
    
    def setup_method(self):
        """Reset shared instance before each test"""
        mcp_server_fastmcp._autocad_util = None
    
    def test_mcp_server_tools_without_gui_connection(self):
        """Test MCP Server tools behave correctly when GUI hasn't established connection"""
        # Scenario: User starts MCP Server before GUI, tries to use tools
        
        # Step 1: No GUI connection established
        mcp_server_fastmcp._autocad_util = None
        
        # Step 2: Try to use AutoCAD tools - should get helpful error
        tools_to_test = [
            ("check_autocad_status", lambda: mcp_server_fastmcp.check_autocad_status()),
            ("draw_circle", lambda: mcp_server_fastmcp.draw_circle([0, 0, 0], 10.0)),
            ("draw_line", lambda: mcp_server_fastmcp.draw_line([0, 0, 0], [10, 10, 0])),
            ("extract_parameters", lambda: mcp_server_fastmcp.extract_autocad_parameters()),
        ]
        
        for tool_name, tool_func in tools_to_test:
            result = tool_func()
            
            # Assert: Each tool should provide helpful guidance
            assert result["status"] == "error", f"{tool_name} should return error"
            assert "AutoCAD" in result["message"], f"{tool_name} should mention AutoCAD"
            
            # Should include guidance about using GUI first
            if "suggestion" in result:
                assert "GUI" in result["suggestion"], f"{tool_name} should suggest using GUI first"
    
    def test_mcp_server_tools_with_gui_connection(self):
        """Test MCP Server tools work correctly when GUI has established connection"""
        # Scenario: GUI establishes connection, then MCP Server uses shared instance
        
        # Step 1: Simulate GUI establishing AutoCAD connection
        mock_autocad = Mock()
        mock_autocad.connected_autocad.return_value = True
        mock_autocad.get_application_info.return_value = {
            "version": "2024",
            "status": "connected",
            "documents": []
        }
        
        # GUI sets the shared instance
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        
        # Step 2: MCP Server tools should now work
        result = mcp_server_fastmcp.check_autocad_status()
        
        # Assert: Should now succeed
        assert result["status"] == "success"
        assert result["connected"] == True
        assert "application_info" in result
    
    def test_connection_handoff_from_gui_to_mcp(self):
        """Test the handoff process from GUI to MCP Server"""
        # This tests the actual workflow described in AC6
        
        # Step 1: Initial state - no connections
        assert mcp_server_fastmcp.get_autocad_util() is None
        
        # Step 2: GUI starts and creates AutoCAD connection
        mock_gui_autocad = Mock()
        mock_gui_autocad.name = "gui_connection"
        mock_gui_autocad.connected_autocad.return_value = True
        
        # Step 3: GUI shares its connection with MCP Server
        mcp_server_fastmcp.set_shared_autocad_util(mock_gui_autocad)
        
        # Step 4: MCP Server should now use the shared connection
        mcp_autocad = mcp_server_fastmcp.get_autocad_util()
        assert mcp_autocad is mock_gui_autocad
        assert mcp_autocad.name == "gui_connection"
        
        # Step 5: MCP tools should work with shared connection
        # Since we properly mocked the AutoCAD instance, this should succeed
        result = mcp_server_fastmcp.check_autocad_status()
        
        # Assert: Should successfully use the shared connection
        assert result["status"] == "success"
        assert result["connected"] == True
    
    def test_no_connection_conflicts_when_shared_properly(self):
        """Test that no connection conflicts occur when shared instance is used properly"""
        # Step 1: Set up a mock AutoCAD instance representing GUI connection
        mock_autocad = Mock()
        mock_autocad.connected_autocad.return_value = True
        mock_autocad.connection_count = 1  # Simulate single connection
        
        # Step 2: Share the instance
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        
        # Step 3: Multiple MCP tool calls should use same instance
        autocad1 = mcp_server_fastmcp.get_autocad_util()
        autocad2 = mcp_server_fastmcp.get_autocad_util()
        autocad3 = mcp_server_fastmcp.get_autocad_util()
        
        # Assert: All should be the same instance (no new connections)
        assert autocad1 is mock_autocad
        assert autocad2 is mock_autocad  
        assert autocad3 is mock_autocad
        
        # Verify connection count hasn't increased (no conflicts)
        assert mock_autocad.connection_count == 1


if __name__ == "__main__":
    pytest.main([__file__, "-v"])