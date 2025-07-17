#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script for direct SSE manager integration
Tests the redesigned MCPSSEManager that uses direct methods
"""

import time
import sys
import os

# Set UTF-8 encoding for Windows
if sys.platform == 'win32':
    os.environ['PYTHONIOENCODING'] = 'utf-8'

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

def test_direct_integration():
    """Test the direct integration SSE manager"""
    print("=== 測試直接整合的 SSE 管理器 ===\n")
    
    # Create manager with custom callback
    manager = MCPSSEManager(port=8083)  # Use different port for testing
    
    def status_callback(message, is_running):
        print(f"   狀態更新: {message} (running={is_running})")
    
    manager.set_status_callback(status_callback)
    
    print("1. 啟動伺服器...")
    success = manager.start_server()
    if success:
        print("[OK] 伺服器啟動成功")
    else:
        print("[FAIL] 伺服器啟動失敗")
        # Try to get more info
        status = manager.get_server_status()
        print(f"   狀態詳情: {status}")
        return
    
    # Wait for server to fully start
    time.sleep(2)
    
    print("\n2. 測試伺服器狀態...")
    status = manager.get_server_status()
    print(f"   運行中: {status['is_running']}")
    print(f"   端口: {status['port']}")
    print(f"   健康檢查: {status['health_check']}")
    if 'server_name' in status:
        print(f"   伺服器名稱: {status['server_name']}")
        print(f"   版本: {status['server_version']}")
        print(f"   活動連接: {status['active_connections']}")
    
    print("\n3. 直接測試工具方法...")
    
    # Test connection
    print("   - 測試連接:")
    result = manager.test_tool_directly("test_connection")
    print(f"     {result}")
    
    # Get server info
    print("   - 獲取伺服器資訊:")
    result = manager.test_tool_directly("get_server_info")
    print(f"     {result}")
    
    # Check AutoCAD status
    print("   - 檢查 AutoCAD 狀態:")
    result = manager.test_tool_directly("check_autocad_status")
    print(f"     {result}")
    
    # Check Odoo status
    print("   - 檢查 Odoo 狀態:")
    result = manager.test_tool_directly("check_odoo_status")
    print(f"     {result}")
    
    print("\n4. 測試 HTTP API...")
    result = manager.test_mcp_connection()
    if result["success"]:
        print(f"[OK] HTTP API 測試成功")
        print(f"  工具數量: {result['tools_count']}")
        print(f"  可用工具: {', '.join(result['tools'])}")
        print(f"  測試結果: {result['test_result']}")
    else:
        print(f"[FAIL] HTTP API 測試失敗: {result['error']}")
    
    print("\n5. 停止伺服器...")
    manager.stop_server()
    print("[OK] 伺服器已停止")
    
    print("\n=== 測試完成 ===")

if __name__ == "__main__":
    try:
        test_direct_integration()
    except Exception as e:
        print(f"錯誤: {e}")
        import traceback
        traceback.print_exc()