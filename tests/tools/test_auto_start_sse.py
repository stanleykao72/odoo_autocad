#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test Auto-Start SSE Server Functionality
Tests that the SSE server starts automatically when GUI launches
"""

import sys
import os
import threading
import time
import requests
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from utility.util_mcp_sse_manager import MCPSSEManager

def test_auto_start_sse_server():
    """Test that SSE server can be started automatically"""
    print("🧪 測試自動啟動 SSE 伺服器功能...")
    
    # Create SSE manager
    sse_manager = MCPSSEManager(port=8085)  # Use different port for testing
    
    # Set up status callback to track status changes
    status_updates = []
    
    def status_callback(message, is_running):
        status_updates.append((message, is_running))
        print(f"📢 SSE 狀態更新: {message} (運行中: {is_running})")
    
    sse_manager.set_status_callback(status_callback)
    
    # Test auto-start simulation
    print("🚀 模擬自動啟動...")
    
    def auto_start_sse_server():
        """Simulate the auto-start functionality"""
        try:
            # This simulates what happens in form_main_modern.py
            threading.Thread(target=sse_manager.start_server, daemon=True).start()
            return True
        except Exception as e:
            print(f"❌ 自動啟動失敗: {e}")
            return False
    
    # Start the server
    success = auto_start_sse_server()
    
    if success:
        print("✅ 自動啟動指令已執行")
        
        # Wait for server to start
        print("⏳ 等待伺服器啟動...")
        time.sleep(3)
        
        # Check server status
        status = sse_manager.get_server_status()
        print(f"📊 伺服器狀態: {status}")
        
        # Test connection
        print("🔗 測試連接...")
        try:
            response = requests.get(
                f"http://localhost:8085/sse",
                timeout=5,
                headers={"Accept": "text/event-stream"}
            )
            if response.status_code == 200:
                print("✅ SSE 端點連接成功")
                print(f"📡 響應狀態: {response.status_code}")
            else:
                print(f"⚠️ 連接異常，狀態碼: {response.status_code}")
        except Exception as e:
            print(f"❌ 連接測試失敗: {e}")
        
        # Clean up
        print("🧹 清理資源...")
        sse_manager.stop_server()
        time.sleep(1)
        
        print("✅ 測試完成")
        
    else:
        print("❌ 自動啟動失敗")
    
    # Print status updates
    print("\n📋 狀態更新記錄:")
    for message, is_running in status_updates:
        print(f"  - {message} (運行中: {is_running})")

if __name__ == "__main__":
    # Set UTF-8 encoding for Windows
    if sys.platform == "win32":
        os.environ["PYTHONIOENCODING"] = "utf-8"
    
    test_auto_start_sse_server()