# -*- coding: utf-8 -*-
"""
Test AC6: COM connection conflict resolution for MCP Server
Tests shared AutoCAD instance functionality and graceful degradation
"""

import pytest
import sys
import os
from unittest.mock import Mock, patch

# Add project root to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

# Import the MCP server module
import mcp_server_fastmcp


class TestAC6SharedInstanceHandling:
    """Test AC6: COM connection conflict resolution"""
    
    def setup_method(self):
        """Reset shared instance before each test"""
        mcp_server_fastmcp._autocad_util = None
    
    def test_get_autocad_util_returns_none_when_no_shared_instance(self):
        """Test that get_autocad_util() returns None when no shared instance (AC6)"""
        # Arrange: No shared instance set
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.get_autocad_util()
        
        # Assert: Should return None (not auto-create)
        assert result is None
    
    def test_get_autocad_util_returns_shared_instance_when_set(self):
        """Test that get_autocad_util() returns shared instance when available"""
        # Arrange: Set a mock shared instance
        mock_autocad = Mock()
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        
        # Act
        result = mcp_server_fastmcp.get_autocad_util()
        
        # Assert: Should return the shared instance
        assert result is mock_autocad
    
    def test_set_shared_autocad_util_stores_instance(self):
        """Test that set_shared_autocad_util() properly stores the instance"""
        # Arrange
        mock_autocad = Mock()
        
        # Act
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        
        # Assert
        assert mcp_server_fastmcp._autocad_util is mock_autocad
        assert mcp_server_fastmcp.get_autocad_util() is mock_autocad
    
    def test_check_autocad_status_graceful_degradation(self):
        """Test check_autocad_status() provides helpful error when no shared instance"""
        # Arrange: No shared instance
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.check_autocad_status()
        
        # Assert: Should provide helpful error message
        assert result["status"] == "error"
        assert result["connected"] == False
        assert "shared AutoCAD instance" in result["error"]
        assert "GUI first" in result["suggestion"]
        assert result["connection_type"] == "shared_instance_required"
    
    def test_draw_circle_graceful_degradation(self):
        """Test draw_circle() provides helpful error when no shared instance"""
        # Arrange: No shared instance
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.draw_circle(
            center_point=[0, 0, 0],
            radius=10.0,
            layer="0"
        )
        
        # Assert: Should provide helpful error message
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "GUI" in result["suggestion"]
        assert result["connection_type"] == "shared_instance_required"
    
    def test_draw_line_graceful_degradation(self):
        """Test draw_line() provides helpful error when no shared instance"""
        # Arrange: No shared instance
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.draw_line(
            start_point=[0, 0, 0],
            end_point=[10, 10, 0],
            layer="0"
        )
        
        # Assert: Should provide helpful error message
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        assert "GUI" in result["suggestion"]
        assert result["connection_type"] == "shared_instance_required"
    
    def test_extract_autocad_parameters_graceful_degradation(self):
        """Test extract_autocad_parameters() provides helpful error when no shared instance"""
        # Arrange: No shared instance
        mcp_server_fastmcp._autocad_util = None
        
        # Act
        result = mcp_server_fastmcp.extract_autocad_parameters()
        
        # Assert: Should provide helpful error message
        assert result["status"] == "error"
        assert "shared AutoCAD instance" in result["error"]
        assert "GUI first" in result["suggestion"]
        assert result["connection_type"] == "shared_instance_required"
    
    def test_multiple_tool_calls_consistent_behavior(self):
        """Test that multiple tool calls consistently handle missing shared instance"""
        # Arrange: No shared instance
        mcp_server_fastmcp._autocad_util = None
        
        # Act: Call multiple tools
        status_result = mcp_server_fastmcp.check_autocad_status()
        circle_result = mcp_server_fastmcp.draw_circle([0, 0, 0], 5.0)
        line_result = mcp_server_fastmcp.draw_line([0, 0, 0], [5, 5, 0])
        
        # Assert: All should provide consistent error response
        for result in [status_result, circle_result, line_result]:
            assert result["status"] == "error"
            assert "AutoCAD" in result["message"]
            # Each should provide guidance about using GUI first
            assert "suggestion" in result
    
    def test_shared_instance_workflow_simulation(self):
        """Test the complete workflow: no instance -> set instance -> tools work"""
        # Step 1: No shared instance - tools should fail gracefully
        mcp_server_fastmcp._autocad_util = None
        
        result_no_instance = mcp_server_fastmcp.check_autocad_status()
        assert result_no_instance["status"] == "error"
        assert result_no_instance["connection_type"] == "shared_instance_required"
        
        # Step 2: Set shared instance - tools should now work
        mock_autocad = Mock()
        mock_autocad.connected_autocad.return_value = True
        mock_autocad.get_application_info.return_value = {
            "version": "2024",
            "status": "connected"
        }
        
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        
        # Step 3: Tools should now work with shared instance
        result_with_instance = mcp_server_fastmcp.check_autocad_status()
        assert result_with_instance["status"] == "success"
        assert result_with_instance["connected"] == True
    
    @patch('mcp_server_fastmcp.logger')
    def test_shared_instance_logging(self, mock_logger):
        """Test that setting shared instance generates appropriate log messages"""
        # Arrange
        mock_autocad = Mock()
        
        # Act
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        
        # Assert: Should log the shared instance setup
        mock_logger.info.assert_called_with("Shared AutoCAD utility instance set from GUI")


class TestConnectionConflictDetection:
    """Test connection conflict detection mechanisms"""
    
    def test_no_duplicate_connection_attempts(self):
        """Test that MCP server doesn't try to create duplicate connections"""
        # This test verifies that get_autocad_util() doesn't auto-create
        # connections when a shared instance should be used
        
        # Arrange: No shared instance
        mcp_server_fastmcp._autocad_util = None
        
        # Act: Multiple calls to get_autocad_util()
        result1 = mcp_server_fastmcp.get_autocad_util()
        result2 = mcp_server_fastmcp.get_autocad_util()
        result3 = mcp_server_fastmcp.get_autocad_util()
        
        # Assert: All should return None (no auto-creation)
        assert result1 is None
        assert result2 is None
        assert result3 is None
    
    def test_shared_instance_priority(self):
        """Test that shared instance takes priority over any potential auto-creation"""
        # Arrange: Set a shared instance
        mock_shared_autocad = Mock()
        mock_shared_autocad.name = "shared_instance"
        mcp_server_fastmcp.set_shared_autocad_util(mock_shared_autocad)
        
        # Act: Get AutoCAD util
        result = mcp_server_fastmcp.get_autocad_util()
        
        # Assert: Should return the shared instance
        assert result is mock_shared_autocad
        assert result.name == "shared_instance"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])