#!/usr/bin/env python
# -*- coding: utf-8 -*-

import time
import requests
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

def simple_test():
    print("=== 簡單測試 ===")
    
    # Create manager
    manager = MCPSSEManager(port=8084)
    
    print("1. 啟動伺服器...")
    success = manager.start_server()
    print(f"   啟動結果: {success}")
    
    # Test direct tool calls
    print("\n2. 直接工具測試...")
    if manager.server:
        print("   伺服器物件存在")
        try:
            result = manager.test_tool_directly("test_connection")
            print(f"   test_connection: {result}")
        except Exception as e:
            print(f"   錯誤: {e}")
    else:
        print("   伺服器物件不存在")
    
    # Test HTTP
    print("\n3. 快速 HTTP 測試...")
    try:
        response = requests.get(f"http://localhost:{manager.port}/sse", timeout=5)
        print(f"   HTTP 狀態碼: {response.status_code}")
        if response.status_code == 200:
            print("   HTTP 連接成功")
        else:
            print("   HTTP 連接失敗")
    except Exception as e:
        print(f"   HTTP 錯誤: {e}")
    
    print("\n4. 清理...")
    manager.cleanup()
    print("   完成")

if __name__ == "__main__":
    simple_test()