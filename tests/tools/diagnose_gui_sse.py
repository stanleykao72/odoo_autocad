#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
診斷 GUI SSE 問題
"""

import time
import requests
import threading
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

def diagnose_gui_sse():
    """診斷 GUI SSE 實際問題"""
    print("=== 診斷 GUI SSE 問題 ===\n")
    
    # 模擬 GUI 環境
    manager = MCPSSEManager(port=8083)
    
    print("1. 檢查初始狀態...")
    print(f"   is_running: {manager.is_running}")
    print(f"   server: {manager.server}")
    print(f"   server_thread: {manager.server_thread}")
    
    # 設置狀態回調
    status_messages = []
    def status_callback(message, is_running):
        status_messages.append((message, is_running))
        print(f"   [STATUS] {message} (running={is_running})")
    
    manager.set_status_callback(status_callback)
    
    print("\n2. 嘗試啟動 SSE 伺服器...")
    success = manager.start_server()
    print(f"   啟動結果: {success}")
    
    # 檢查啟動後狀態
    print(f"   is_running: {manager.is_running}")
    print(f"   server: {manager.server}")
    print(f"   server_thread: {manager.server_thread}")
    print(f"   server_thread.is_alive(): {manager.server_thread.is_alive() if manager.server_thread else 'N/A'}")
    
    # 等待啟動
    print("\n3. 等待 5 秒讓伺服器完全啟動...")
    time.sleep(5)
    
    # 檢查健康狀態
    print("\n4. 檢查健康狀態...")
    health = manager._check_server_health()
    print(f"   健康檢查: {health}")
    
    # 測試端口連接
    print("\n5. 直接測試端口連接...")
    try:
        response = requests.get(f"http://localhost:8083/sse", timeout=5)
        print(f"   HTTP 狀態碼: {response.status_code}")
        print(f"   響應頭: {dict(response.headers)}")
        if response.status_code == 200:
            print("   [OK] 端口連接成功")
        else:
            print("   [FAIL] 端口連接失敗")
    except Exception as e:
        print(f"   [ERROR] 端口連接錯誤: {e}")
    
    # 測試 MCP JSON-RPC
    print("\n6. 測試 MCP JSON-RPC...")
    try:
        response = requests.post(
            "http://localhost:8083/sse",
            json={
                "jsonrpc": "2.0",
                "id": 1,
                "method": "tools/list",
                "params": {}
            },
            timeout=5
        )
        if response.status_code == 200:
            data = response.json()
            tools = data.get('result', {}).get('tools', [])
            print(f"   [OK] MCP 測試成功，工具數量: {len(tools)}")
            for tool in tools:
                print(f"     - {tool['name']}")
        else:
            print(f"   [FAIL] MCP 測試失敗: {response.status_code}")
    except Exception as e:
        print(f"   [ERROR] MCP 測試錯誤: {e}")
    
    # 檢查伺服器內部狀態
    print("\n7. 檢查伺服器內部狀態...")
    if manager.server:
        print(f"   伺服器名稱: {manager.server.name}")
        print(f"   伺服器版本: {manager.server.version}")
        print(f"   活動連接: {len(manager.server.active_connections)}")
        print(f"   FastAPI app: {manager.server.app}")
    else:
        print("   [ERROR] 伺服器物件為 None")
    
    # 檢查線程狀態
    print("\n8. 檢查線程詳細狀態...")
    if manager.server_thread:
        print(f"   線程 alive: {manager.server_thread.is_alive()}")
        print(f"   線程 daemon: {manager.server_thread.daemon}")
        print(f"   線程 name: {manager.server_thread.name}")
    else:
        print("   [ERROR] 伺服器線程為 None")
    
    # 檢查事件循環
    print("\n9. 檢查事件循環狀態...")
    if hasattr(manager, '_loop') and manager._loop:
        print(f"   事件循環: {manager._loop}")
        print(f"   事件循環運行中: {manager._loop.is_running()}")
        print(f"   事件循環關閉: {manager._loop.is_closed()}")
    else:
        print("   [INFO] 事件循環未設置或已清理")
    
    # 狀態訊息摘要
    print("\n10. 狀態訊息摘要...")
    for i, (msg, running) in enumerate(status_messages):
        print(f"    {i+1}. {msg} (running={running})")
    
    print("\n11. 保持運行 10 秒供外部測試...")
    for i in range(10, 0, -1):
        print(f"    剩餘 {i} 秒...", end='\r')
        time.sleep(1)
    
    print("\n\n12. 停止伺服器...")
    manager.stop_server()
    print("    伺服器已停止")
    
    print("\n=== 診斷完成 ===")

if __name__ == "__main__":
    diagnose_gui_sse()