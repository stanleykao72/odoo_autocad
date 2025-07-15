#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MCP連接測試腳本

用於測試 MCP 伺服器的基本連接和功能
"""

import socket
import json
import time
import sys

def test_tcp_connection(host='localhost', port=8000, timeout=5):
    """測試TCP連接到MCP伺服器"""
    print(f"🔍 測試TCP連接到 {host}:{port}")
    
    try:
        # 建立socket連接
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        sock.connect((host, port))
        
        print("✅ TCP連接成功")
        
        # 發送tools/list請求
        tools_list_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/list",
            "params": {}
        }
        
        request_data = json.dumps(tools_list_request).encode('utf-8')
        sock.send(request_data)
        
        # 接收回應
        response_data = sock.recv(4096)
        response = json.loads(response_data.decode('utf-8'))
        
        print("✅ 成功獲取工具列表")
        print(f"📋 可用工具數量: {len(response.get('result', {}).get('tools', []))}")
        
        # 顯示工具詳情
        tools = response.get('result', {}).get('tools', [])
        for tool in tools:
            print(f"   🔧 {tool['name']}: {tool['description']}")
        
        sock.close()
        return True
        
    except socket.timeout:
        print(f"❌ 連接超時 ({timeout}秒)")
        return False
    except ConnectionRefusedError:
        print("❌ 連接被拒絕 - 伺服器可能未啟動")
        return False
    except Exception as e:
        print(f"❌ 連接錯誤: {e}")
        return False

def test_tool_execution(host='localhost', port=8000):
    """測試工具執行"""
    print(f"\n🔍 測試工具執行")
    
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.connect((host, port))
        
        # 測試scan_all_entities工具
        tool_call_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/call",
            "params": {
                "name": "scan_all_entities",
                "arguments": {}
            }
        }
        
        request_data = json.dumps(tool_call_request).encode('utf-8')
        sock.send(request_data)
        
        # 接收回應
        response_data = sock.recv(8192)
        response = json.loads(response_data.decode('utf-8'))
        
        if 'result' in response:
            print("✅ 工具執行成功")
            content = response['result']['content'][0]['text']
            result = json.loads(content)
            print(f"📊 掃描結果: {result.get('success', False)}")
            if result.get('entities'):
                print(f"🎯 發現實體數量: {len(result['entities'])}")
        else:
            print(f"❌ 工具執行失敗: {response.get('error', {}).get('message', 'Unknown error')}")
        
        sock.close()
        return True
        
    except Exception as e:
        print(f"❌ 工具執行錯誤: {e}")
        return False

def main():
    """主測試函數"""
    print("🚀 AutoCAD-Odoo MCP 伺服器連接測試")
    print("=" * 50)
    
    # 檢查命令列參數
    host = 'localhost'
    port = 8000
    
    if len(sys.argv) > 1:
        port = int(sys.argv[1])
    if len(sys.argv) > 2:
        host = sys.argv[2]
    
    print(f"🎯 測試目標: {host}:{port}")
    print()
    
    # 測試基本連接
    if not test_tcp_connection(host, port):
        print("\n❌ 基本連接測試失敗")
        print("\n💡 故障排除建議:")
        print("1. 確認MCP伺服器已啟動")
        print("2. 檢查端口是否正確")
        print("3. 確認防火牆設定")
        print("\n啟動伺服器指令:")
        print(f"   python odoo.py --mcp-server --mcp-port {port}")
        return 1
    
    # 測試工具執行
    print()
    if test_tool_execution(host, port):
        print("\n✅ 所有測試通過！MCP伺服器運行正常")
    else:
        print("\n⚠️  基本連接成功，但工具執行有問題")
        print("    可能原因：AutoCAD未連接或Odoo配置問題")
    
    print("\n🎉 測試完成")
    return 0

if __name__ == '__main__':
    sys.exit(main())