# -*- coding: utf-8 -*-
"""
Test MCP Protocol Compliance
測試 MCP 協定相容性
"""

import sys
import os
import requests
import json
import time
from pathlib import Path

# Set UTF-8 encoding for Windows console
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

class MCPProtocolTester:
    """MCP 協定測試器"""
    
    def __init__(self):
        self.base_url = "http://localhost:8000"
    
    def test_mcp_handshake(self):
        """測試 MCP 握手協定"""
        print("🤝 測試 MCP 握手協定...")
        
        # Standard MCP initialization message
        init_message = {
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
                    "name": "gemini-cli",
                    "version": "1.0.0"
                }
            }
        }
        
        try:
            # Test JSON-RPC over HTTP POST
            response = requests.post(
                f"{self.base_url}/mcp",
                json=init_message,
                headers={'Content-Type': 'application/json'},
                timeout=10
            )
            
            print(f"  📡 JSON-RPC POST /mcp: HTTP {response.status_code}")
            print(f"  📋 Response: {response.text[:200]}")
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    if "result" in data:
                        print("  ✅ Valid MCP JSON-RPC response!")
                        return True
                except:
                    pass
                    
        except Exception as e:
            print(f"  ❌ JSON-RPC test failed: {e}")
        
        # Test if it's stdio-based MCP instead
        try:
            response = requests.get(f"{self.base_url}/", timeout=5)
            print(f"  📡 HTTP GET /: {response.status_code}")
            
            if "mcp" in response.text.lower() or "stdio" in response.text.lower():
                print("  ⚠️  可能是 STDIO-based MCP，不支援 HTTP")
                return False
                
        except Exception as e:
            print(f"  ❌ HTTP test failed: {e}")
        
        return False
    
    def test_sse_format(self):
        """測試 SSE 格式是否符合 MCP 規範"""
        print("\n🌊 測試 SSE 格式合規性...")
        
        try:
            headers = {
                'Accept': 'text/event-stream',
                'Cache-Control': 'no-cache'
            }
            
            response = requests.get(
                f"{self.base_url}/sse", 
                headers=headers, 
                stream=True, 
                timeout=10
            )
            
            print(f"  📡 SSE Connection: HTTP {response.status_code}")
            print(f"  📋 Headers: {dict(response.headers)}")
            
            if response.status_code == 200:
                print("  📺 Reading SSE events...")
                
                events = []
                start_time = time.time()
                
                for line in response.iter_lines(decode_unicode=True):
                    if time.time() - start_time > 5:  # 5 second timeout
                        break
                    
                    if line:
                        events.append(line)
                        print(f"    📨 {line}")
                        
                        # Check for MCP-style messages
                        if line.startswith("data:"):
                            try:
                                data_part = line[5:].strip()  # Remove "data: "
                                if data_part:
                                    json_data = json.loads(data_part)
                                    if "jsonrpc" in json_data:
                                        print("      ✅ Found JSON-RPC in SSE data!")
                                        return True
                            except:
                                pass
                
                print(f"  📊 Total events received: {len(events)}")
                
                # Check if it looks like MCP SSE format
                if any("data:" in event for event in events):
                    print("  ⚠️  SSE format detected, but no MCP JSON-RPC found")
                    return False
                    
        except Exception as e:
            print(f"  ❌ SSE test failed: {e}")
        
        return False
    
    def test_transport_methods(self):
        """測試不同的傳輸方法"""
        print("\n🚀 測試不同傳輸方法...")
        
        # Method 1: WebSocket
        try:
            import websocket
            ws_url = self.base_url.replace('http:', 'ws:') + '/ws'
            print(f"  🔌 Testing WebSocket: {ws_url}")
            
            ws = websocket.create_connection(ws_url, timeout=3)
            ws.close()
            print("  ✅ WebSocket connection successful")
        except Exception as e:
            print(f"  ❌ WebSocket failed: {str(e)[:100]}")
        
        # Method 2: Different SSE endpoints
        sse_endpoints = ['/sse', '/events', '/stream', '/mcp-sse']
        
        for endpoint in sse_endpoints:
            try:
                url = f"{self.base_url}{endpoint}"
                response = requests.get(url, timeout=3)
                print(f"  📡 {endpoint}: HTTP {response.status_code}")
                
                if response.status_code == 200:
                    content_type = response.headers.get('content-type', '')
                    if 'text/event-stream' in content_type:
                        print(f"    ✅ Valid SSE endpoint found!")
                        
            except Exception as e:
                print(f"  ❌ {endpoint}: {str(e)[:50]}")
    
    def generate_alternative_configs(self):
        """生成替代的 Gemini 配置選項"""
        print("\n📝 生成替代配置選項...")
        
        configs = [
            # Standard MCP over SSE
            {
                "transport": "sse",
                "url": "http://localhost:8000/sse"
            },
            # Standard MCP over HTTP
            {
                "transport": "http", 
                "url": "http://localhost:8000"
            },
            # STDIO-based (if supported)
            {
                "command": "python",
                "args": [str(Path(__file__).parent.parent / "odoo.py"), "--mcp-server"],
                "transport": "stdio"
            },
            # WebSocket (if supported)
            {
                "transport": "websocket",
                "url": "ws://localhost:8000/ws"
            }
        ]
        
        print("\n🔧 Gemini CLI 配置選項:")
        for i, config in enumerate(configs, 1):
            print(f"\n選項 {i}:")
            print(json.dumps({"autocad": config}, indent=2, ensure_ascii=False))
        
        return configs
    
    def run_protocol_test(self):
        """執行完整協定測試"""
        print("🔍 MCP Protocol Compliance Test")
        print("=" * 80)
        
        # Test 1: MCP handshake
        mcp_works = self.test_mcp_handshake()
        
        # Test 2: SSE format
        sse_works = self.test_sse_format()
        
        # Test 3: Transport methods
        self.test_transport_methods()
        
        # Test 4: Generate configs
        configs = self.generate_alternative_configs()
        
        print("\n" + "=" * 80)
        print("Protocol Test Results")
        print("=" * 80)
        
        if mcp_works:
            print("✅ MCP JSON-RPC Protocol: Working")
            print("💡 建議使用標準 MCP over HTTP 配置")
        elif sse_works:
            print("✅ SSE Transport: Working")
            print("⚠️  但可能不是標準 MCP 格式")
            print("💡 嘗試 SSE 配置或 STDIO 配置")
        else:
            print("❌ Standard MCP Protocol: Not detected")
            print("💡 可能需要使用 STDIO transport")
        
        # Specific recommendation
        print(f"\n🎯 推薦配置:")
        if mcp_works:
            recommended = configs[1]  # HTTP transport
        else:
            recommended = configs[2]  # STDIO transport
        
        print(json.dumps({"autocad": recommended}, indent=2, ensure_ascii=False))
        
        return mcp_works or sse_works

def main():
    """主程式"""
    tester = MCPProtocolTester()
    
    try:
        success = tester.run_protocol_test()
        
        if success:
            print("\n🎉 找到相容的 MCP 協定配置！")
        else:
            print("\n⚠️  未能找到標準 MCP 協定，建議使用 STDIO transport")
        
        return 0 if success else 1
        
    except KeyboardInterrupt:
        print("\n\n⚠️  測試被用戶中斷")
        return 1

if __name__ == "__main__":
    sys.exit(main())