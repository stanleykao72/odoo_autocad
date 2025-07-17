#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
驗證 SSE 伺服器可以從 GUI 啟動並與 Gemini CLI 連接
"""

import time
import requests
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

def verify_sse_integration():
    """驗證 SSE 整合"""
    print("=== 驗證 GUI-SSE-Gemini CLI 整合 ===\n")
    
    # 步驟 1: 從 GUI 啟動 SSE 伺服器
    print("1. 模擬從 GUI 啟動 SSE 伺服器...")
    manager = MCPSSEManager(port=8083)
    
    def status_callback(message, is_running):
        print(f"   [GUI] {message}")
    
    manager.set_status_callback(status_callback)
    
    success = manager.start_server()
    if not success:
        print("   [FAIL] 無法啟動 SSE 伺服器")
        return
    
    # 等待伺服器完全啟動
    time.sleep(3)
    
    # 步驟 2: 驗證 SSE 伺服器狀態
    print("\n2. 驗證伺服器狀態...")
    status = manager.get_server_status()
    print(f"   運行狀態: {status['is_running']}")
    print(f"   端口: {status['port']}")
    print(f"   健康檢查: {status['health_check']}")
    
    # 步驟 3: 測試 MCP JSON-RPC 協議
    print("\n3. 測試 MCP JSON-RPC 協議...")
    
    # Initialize
    init_request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "clientInfo": {
                "name": "Gemini CLI",
                "version": "1.0.0"
            }
        }
    }
    
    try:
        response = requests.post(
            f"http://localhost:8083/sse",
            json=init_request,
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            print("   [OK] Initialize 成功")
            print(f"   協議版本: {data['result']['protocolVersion']}")
            print(f"   伺服器: {data['result']['serverInfo']['name']}")
        else:
            print(f"   [FAIL] Initialize 失敗: {response.status_code}")
            
    except Exception as e:
        print(f"   [FAIL] Initialize 錯誤: {e}")
    
    # Tools list
    tools_request = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list",
        "params": {}
    }
    
    try:
        response = requests.post(
            f"http://localhost:8083/sse",
            json=tools_request,
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            tools = data['result']['tools']
            print(f"   [OK] Tools list 成功 ({len(tools)} 個工具)")
            for tool in tools:
                print(f"     - {tool['name']}: {tool['description']}")
        else:
            print(f"   [FAIL] Tools list 失敗: {response.status_code}")
            
    except Exception as e:
        print(f"   [FAIL] Tools list 錯誤: {e}")
    
    # Tool call
    call_request = {
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "test_connection",
            "arguments": {}
        }
    }
    
    try:
        response = requests.post(
            f"http://localhost:8083/sse",
            json=call_request,
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            result = data['result']
            print(f"   [OK] Tool call 成功")
            print(f"     結果: {result}")
        else:
            print(f"   [FAIL] Tool call 失敗: {response.status_code}")
            
    except Exception as e:
        print(f"   [FAIL] Tool call 錯誤: {e}")
    
    # 步驟 4: 顯示 Gemini CLI 配置
    print("\n4. Gemini CLI 配置驗證...")
    print("   當前配置應為:")
    print(f"""   {{
     "autocad-odoo-sse": {{
       "url": "http://localhost:8083/sse",
       "timeout": 30000,
       "description": "AutoCAD-Odoo Integration with Official SSE"
     }}
   }}""")
    
    print("\n5. 使用說明:")
    print("   - 重新啟動 Gemini CLI")
    print("   - 檢查連接狀態應顯示: 🟢 autocad-odoo-sse - Ready (4 tools)")
    print("   - 可以使用指令: \"測試 MCP 連接\" 或 \"獲取伺服器資訊\"")
    
    # 保持運行
    print("\n6. 保持伺服器運行 20 秒以供 Gemini CLI 連接...")
    print("   請在這段時間內啟動 Gemini CLI 進行測試")
    
    for i in range(20, 0, -1):
        print(f"   剩餘時間: {i} 秒", end='\r')
        time.sleep(1)
    
    print("\n\n7. 停止伺服器...")
    manager.stop_server()
    print("   [OK] 伺服器已停止")
    
    print("\n=== 驗證完成 ===")
    print("如果所有測試都通過，表示 GUI-SSE-Gemini CLI 整合成功！")

if __name__ == "__main__":
    verify_sse_integration()