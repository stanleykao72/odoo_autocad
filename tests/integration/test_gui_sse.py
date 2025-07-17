#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test GUI SSE integration
"""

import time
import threading
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

def test_gui_sse():
    """Test GUI SSE integration on port 8082"""
    print("=== GUI SSE 整合測試 ===")
    
    # 模擬 GUI 使用方式
    manager = MCPSSEManager(port=8083)
    
    def status_callback(message, is_running):
        print(f"[STATUS] {message} (運行: {is_running})")
    
    manager.set_status_callback(status_callback)
    
    print("1. 啟動 SSE 伺服器...")
    success = manager.start_server()
    
    if success:
        print("   [OK] 啟動成功")
        
        # 等待伺服器完全啟動
        time.sleep(3)
        
        print("2. 檢查伺服器狀態...")
        status = manager.get_server_status()
        print(f"   運行狀態: {status['is_running']}")
        print(f"   端口: {status['port']}")
        print(f"   健康檢查: {status['health_check']}")
        
        if 'server_name' in status:
            print(f"   伺服器名稱: {status['server_name']}")
            print(f"   版本: {status['server_version']}")
            print(f"   活動連接: {status['active_connections']}")
        
        print("3. 測試直接工具調用...")
        test_result = manager.test_tool_directly("test_connection")
        print(f"   test_connection: {test_result}")
        
        print("4. 測試 HTTP API...")
        api_result = manager.test_mcp_connection()
        if api_result["success"]:
            print("   [OK] HTTP API 測試成功")
            print(f"   工具數量: {api_result['tools_count']}")
            print(f"   可用工具: {', '.join(api_result['tools'])}")
        else:
            print(f"   [FAIL] HTTP API 測試失敗: {api_result['error']}")
        
        print("5. 保持運行 10 秒...")
        time.sleep(10)
        
        print("6. 停止伺服器...")
        manager.stop_server()
        print("   [OK] 已停止")
        
    else:
        print("   [FAIL] 啟動失敗")
    
    print("\n=== 測試完成 ===")

if __name__ == "__main__":
    test_gui_sse()