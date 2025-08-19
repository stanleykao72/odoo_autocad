#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Manual test for AC6: COM connection conflict resolution
Tests the shared AutoCAD instance functionality with the running GUI
"""

import sys
import os
import requests
import json
import time

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

def test_mcp_server_health():
    """Test if MCP server is running and healthy"""
    try:
        response = requests.get("http://localhost:8084/health", timeout=5)
        if response.status_code == 200:
            health_data = response.json()
            print("[PASS] MCP Server Health Check - PASSED")
            print(f"   Server: {health_data.get('server', {}).get('name', 'Unknown')}")
            print(f"   Version: {health_data.get('server', {}).get('version', 'Unknown')}")
            print(f"   Status: {health_data.get('status', 'Unknown')}")
            return True
        else:
            print(f"[FAIL] MCP Server Health Check - FAILED (HTTP {response.status_code})")
            return False
    except Exception as e:
        print(f"[FAIL] MCP Server Health Check - FAILED (Connection Error: {e})")
        return False

def test_shared_autocad_instance():
    """Test shared AutoCAD instance functionality by importing the MCP server directly"""
    try:
        # Import the MCP server to test shared instance functionality
        import mcp_server_fastmcp
        
        print("\n=== Testing AC6: Shared AutoCAD Instance ===")
        
        # Test 1: No shared instance (should return None)
        mcp_server_fastmcp._autocad_util = None
        autocad_util = mcp_server_fastmcp.get_autocad_util()
        
        if autocad_util is None:
            print("[PASS] Test 1 - get_autocad_util() returns None when no shared instance - PASSED")
        else:
            print("[FAIL] Test 1 - get_autocad_util() should return None - FAILED")
            return False
        
        # Test 2: check_autocad_status graceful degradation
        result = mcp_server_fastmcp.check_autocad_status()
        
        if (result["status"] == "error" and 
            "shared AutoCAD instance" in result["error"] and
            "connection_type" in result and
            result["connection_type"] == "shared_instance_required"):
            print("[PASS] Test 2 - check_autocad_status graceful degradation - PASSED")
            print(f"   Error message: {result['message']}")
            print(f"   Suggestion: {result.get('suggestion', 'None')}")
        else:
            print("[FAIL] Test 2 - check_autocad_status graceful degradation - FAILED")
            print(f"   Result: {result}")
            return False
        
        # Test 3: Drawing tools graceful degradation
        circle_result = mcp_server_fastmcp.draw_circle([0, 0, 0], 10.0)
        
        if (circle_result["status"] == "error" and
            circle_result["error_code"] == "AUTOCAD_NOT_CONNECTED" and
            "connection_type" in circle_result and
            circle_result["connection_type"] == "shared_instance_required"):
            print("[PASS] Test 3 - draw_circle graceful degradation - PASSED")
        else:
            print("[FAIL] Test 3 - draw_circle graceful degradation - FAILED")
            print(f"   Result: {circle_result}")
            return False
        
        # Test 4: Test shared instance setting
        from unittest.mock import Mock
        mock_autocad = Mock()
        mock_autocad.name = "test_shared_instance"
        
        mcp_server_fastmcp.set_shared_autocad_util(mock_autocad)
        retrieved_util = mcp_server_fastmcp.get_autocad_util()
        
        if retrieved_util is mock_autocad and retrieved_util.name == "test_shared_instance":
            print("[PASS] Test 4 - Shared instance setting and retrieval - PASSED")
        else:
            print("[FAIL] Test 4 - Shared instance setting and retrieval - FAILED")
            return False
        
        print("[PASS] All AC6 Tests - PASSED")
        return True
        
    except Exception as e:
        print(f"[FAIL] AC6 Test Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_gui_mcp_integration():
    """Test if GUI has started MCP server correctly"""
    print("\n=== Testing GUI-MCP Integration ===")
    
    # Check if port 8084 is in use (GUI MCP server)
    try:
        response = requests.get("http://localhost:8084/", timeout=5)
        if response.status_code == 200:
            server_info = response.json()
            print("[PASS] GUI MCP Server - RUNNING")
            print(f"   Server: {server_info.get('name', 'Unknown')}")
            print(f"   Version: {server_info.get('version', 'Unknown')}")
            print(f"   Protocol: {server_info.get('protocol', 'Unknown')}")
            print(f"   Transport: {server_info.get('transport', [])}")
            return True
        else:
            print("[FAIL] GUI MCP Server - Not responding correctly")
            return False
    except Exception as e:
        print(f"[FAIL] GUI MCP Server - Connection failed: {e}")
        return False

def main():
    """Run AC6 implementation tests"""
    print("=== AC6: COM Connection Conflict Resolution - Manual Test ===")
    print("Testing the shared AutoCAD instance functionality")
    print()
    
    # Test 1: MCP Server Health
    health_ok = test_mcp_server_health()
    
    # Test 2: GUI-MCP Integration
    gui_mcp_ok = test_gui_mcp_integration()
    
    # Test 3: AC6 Shared Instance Functionality
    ac6_ok = test_shared_autocad_instance()
    
    # Summary
    print("\n" + "="*50)
    print("TEST SUMMARY:")
    print(f"MCP Server Health:     {'[PASS] PASS' if health_ok else '[FAIL] FAIL'}")
    print(f"GUI-MCP Integration:   {'[PASS] PASS' if gui_mcp_ok else '[FAIL] FAIL'}")
    print(f"AC6 Shared Instance:   {'[PASS] PASS' if ac6_ok else '[FAIL] FAIL'}")
    
    overall_result = health_ok and gui_mcp_ok and ac6_ok
    print(f"\nOVERALL RESULT: {'[PASS] ALL TESTS PASSED' if overall_result else '[FAIL] SOME TESTS FAILED'}")
    
    if overall_result:
        print("\n[SUCCESS] AC6 implementation is working correctly!")
        print("   • MCP Server responds to health checks")
        print("   • GUI successfully started MCP server on port 8084")
        print("   • Shared AutoCAD instance functionality works as expected")
        print("   • Graceful degradation messages are helpful and accurate")
    else:
        print("\n[WARNING]  Some issues were detected. Please review the test results above.")
    
    return overall_result

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)