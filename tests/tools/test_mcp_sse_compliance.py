#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
測試 MCP SSE 合規性
檢查是否符合 Gemini CLI 的期望
"""

import json
import requests
import time
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utility.util_mcp_sse_manager import MCPSSEManager

def test_mcp_sse_compliance():
    """測試 MCP SSE 是否符合 Gemini CLI 規範"""
    print("=== 測試 MCP SSE 合規性 ===\n")
    
    # 啟動伺服器
    manager = MCPSSEManager(port=8083)
    
    def status_callback(message, is_running):
        print(f"[SSE] {message}")
    
    manager.set_status_callback(status_callback)
    
    print("1. 啟動 SSE 伺服器...")
    success = manager.start_server()
    if not success:
        print("[FAIL] 無法啟動伺服器")
        return
    
    time.sleep(3)  # 等待啟動
    
    print("\n2. 測試 MCP 初始化握手...")
    
    # 測試 initialize 方法
    init_request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "roots": {
                    "listChanged": True
                },
                "sampling": {}
            },
            "clientInfo": {
                "name": "Gemini CLI",
                "version": "1.0.0"
            }
        }
    }
    
    try:
        response = requests.post(
            "http://localhost:8083/sse",
            json=init_request,
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            print("[OK] Initialize 成功")
            print(f"    協議版本: {data.get('result', {}).get('protocolVersion', 'N/A')}")
            print(f"    伺服器名稱: {data.get('result', {}).get('serverInfo', {}).get('name', 'N/A')}")
            print(f"    伺服器版本: {data.get('result', {}).get('serverInfo', {}).get('version', 'N/A')}")
            
            # 檢查能力
            capabilities = data.get('result', {}).get('capabilities', {})
            print(f"    支援的能力: {list(capabilities.keys())}")
        else:
            print(f"[FAIL] Initialize 失敗: HTTP {response.status_code}")
            print(f"    響應: {response.text}")
            
    except Exception as e:
        print(f"[ERROR] Initialize 錯誤: {e}")
    
    print("\n3. 測試工具列表...")
    
    tools_request = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list",
        "params": {}
    }
    
    try:
        response = requests.post(
            "http://localhost:8083/sse",
            json=tools_request,
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            tools = data.get('result', {}).get('tools', [])
            print(f"[OK] Tools/list 成功，工具數量: {len(tools)}")
            for tool in tools:
                print(f"    - {tool['name']}: {tool['description']}")
        else:
            print(f"[FAIL] Tools/list 失敗: HTTP {response.status_code}")
            
    except Exception as e:
        print(f"[ERROR] Tools/list 錯誤: {e}")
    
    print("\n4. 測試工具調用...")
    
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
            "http://localhost:8083/sse",
            json=call_request,
            headers={"Content-Type": "application/json"},
            timeout=10
        )
        
        if response.status_code == 200:
            data = response.json()
            result = data.get('result')
            print(f"[OK] Tools/call 成功")
            print(f"    結果: {result}")
        else:
            print(f"[FAIL] Tools/call 失敗: HTTP {response.status_code}")
            
    except Exception as e:
        print(f"[ERROR] Tools/call 錯誤: {e}")
    
    print("\n5. 測試 SSE 串流端點...")
    
    try:
        # 測試 SSE 串流
        response = requests.get(
            "http://localhost:8083/sse",
            headers={"Accept": "text/event-stream"},
            stream=True,
            timeout=5
        )
        
        if response.status_code == 200:
            print("[OK] SSE 串流端點可用")
            print(f"    Content-Type: {response.headers.get('content-type', 'N/A')}")
            
            # 讀取前幾個事件
            lines = []
            for i, line in enumerate(response.iter_lines(decode_unicode=True)):
                if i >= 10:  # 只讀取前 10 行
                    break
                if line:
                    lines.append(line)
            
            print(f"    前幾個事件: {lines[:3]}")
        else:
            print(f"[FAIL] SSE 串流失敗: HTTP {response.status_code}")
            
    except Exception as e:
        print(f"[ERROR] SSE 串流錯誤: {e}")
    
    print("\n6. 檢查 CORS 標頭...")
    
    try:
        # OPTIONS 請求檢查 CORS
        response = requests.options("http://localhost:8083/sse", timeout=5)
        
        print(f"    OPTIONS 狀態碼: {response.status_code}")
        cors_headers = {
            k: v for k, v in response.headers.items() 
            if k.lower().startswith('access-control')
        }
        print(f"    CORS 標頭: {cors_headers}")
        
    except Exception as e:
        print(f"[ERROR] CORS 檢查錯誤: {e}")
    
    print("\n7. 模擬 Gemini CLI 連接...")
    print("    Gemini CLI 配置應為:")
    print(f'''    {{
      "autocad-odoo-sse": {{
        "url": "http://localhost:8083/sse",
        "timeout": 30000,
        "description": "AutoCAD-Odoo Integration with SSE"
      }}
    }}''')
    
    print("\n8. 伺服器將保持運行 15 秒...")
    print("    請在這段時間內:")
    print("    1. 重新啟動 Gemini CLI")
    print("    2. 運行 /mcp list")
    print("    3. 檢查 autocad-odoo-sse 是否連接成功")
    
    for i in range(15, 0, -1):
        print(f"    剩餘時間: {i:2d} 秒", end='\r')
        time.sleep(1)
    
    print("\n\n9. 停止伺服器...")
    manager.stop_server()
    print("    伺服器已停止")
    
    print("\n=== 測試完成 ===")
    print("\n結論：")
    print("如果上述所有測試都通過，但 Gemini CLI 仍然無法連接，")
    print("可能的原因包括：")
    print("1. Gemini CLI 對 SSE 傳輸的實現有特殊要求")
    print("2. 需要特定的響應標頭或格式")
    print("3. SSE 傳輸在 MCP 2024-11-05 中已棄用，Gemini CLI 可能不支援")
    print("4. 需要在伺服器啟動後等待更長時間")

if __name__ == "__main__":
    test_mcp_sse_compliance()