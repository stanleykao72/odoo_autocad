#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Simple SSE client for testing MCP SSE Server
"""

import sys
import os
import json
import requests
import threading
import time
from sseclient import SSEClient

# Add the project root to sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class SSETestClient:
    """Simple SSE client for testing"""
    
    def __init__(self, base_url="http://localhost:8080"):
        self.base_url = base_url
        self.post_url = f"{base_url}/sse"
        self.stream_url = f"{base_url}/sse"
        self.session = requests.Session()
    
    def send_mcp_request(self, request_data):
        """Send MCP request via HTTP POST"""
        try:
            response = self.session.post(self.post_url, json=request_data, timeout=5)
            if response.status_code == 200:
                return response.json()
            else:
                print(f"Request failed with status {response.status_code}")
                return None
        except Exception as e:
            print(f"Error sending request: {e}")
            return None
    
    def listen_sse_stream(self, duration=5):
        """Listen to SSE stream for a specified duration"""
        try:
            print(f"🔄 Connecting to SSE stream: {self.stream_url}")
            
            response = self.session.get(self.stream_url, stream=True, timeout=duration+1)
            
            if response.status_code == 200:
                print("✓ Connected to SSE stream")
                
                # Parse SSE events
                events_received = 0
                start_time = time.time()
                
                for line in response.iter_lines(decode_unicode=True):
                    if time.time() - start_time > duration:
                        break
                    
                    if line and line.startswith('data:'):
                        try:
                            data = json.loads(line[5:])  # Remove 'data:' prefix
                            events_received += 1
                            print(f"📡 SSE Event #{events_received}: {data}")
                        except json.JSONDecodeError:
                            print(f"📡 SSE Event #{events_received}: {line[5:]}")
                
                print(f"✓ Received {events_received} SSE events")
                return events_received > 0
            else:
                print(f"✗ SSE connection failed with status {response.status_code}")
                return False
                
        except Exception as e:
            print(f"✗ Error listening to SSE stream: {e}")
            return False
    
    def test_full_mcp_workflow(self):
        """Test complete MCP workflow with SSE"""
        print("\n=== Testing Full MCP Workflow with SSE ===")
        
        # 1. Initialize
        print("\n1. Initializing MCP connection...")
        init_response = self.send_mcp_request({
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "clientInfo": {
                    "name": "SSE Test Client",
                    "version": "1.0.0"
                },
                "capabilities": {
                    "roots": {"listChanged": True},
                    "sampling": {}
                }
            }
        })
        
        if init_response and init_response.get('result'):
            print("✓ Initialization successful")
        else:
            print("✗ Initialization failed")
            return False
        
        # 2. List tools
        print("\n2. Getting tools list...")
        tools_response = self.send_mcp_request({
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list"
        })
        
        if tools_response and tools_response.get('result', {}).get('tools'):
            tools = tools_response['result']['tools']
            print(f"✓ Found {len(tools)} tools:")
            for tool in tools:
                print(f"  - {tool['name']}: {tool['description']}")
        else:
            print("✗ Failed to get tools list")
            return False
        
        # 3. Call a tool
        print("\n3. Calling test_connection tool...")
        call_response = self.send_mcp_request({
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "test_connection",
                "arguments": {}
            }
        })
        
        if call_response and call_response.get('result'):
            print(f"✓ Tool call successful: {call_response['result']}")
        else:
            print("✗ Tool call failed")
            return False
        
        # 4. Test SSE streaming
        print("\n4. Testing SSE streaming...")
        sse_success = self.listen_sse_stream(duration=3)
        
        if sse_success:
            print("✓ SSE streaming test successful")
        else:
            print("✗ SSE streaming test failed")
            return False
        
        return True
    
    def run_tests(self):
        """Run all SSE client tests"""
        print("🧪 Starting SSE Client Tests...")
        
        success = self.test_full_mcp_workflow()
        
        if success:
            print("\n🎉 All SSE client tests passed!")
            print("   MCP SSE Server is working correctly")
            return True
        else:
            print("\n❌ Some SSE client tests failed!")
            return False


def main():
    """Main function"""
    if len(sys.argv) > 1:
        base_url = sys.argv[1]
    else:
        base_url = "http://localhost:8080"
    
    print(f"🔗 Testing SSE server at: {base_url}")
    
    client = SSETestClient(base_url)
    success = client.run_tests()
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    # Set UTF-8 encoding for Windows console
    if sys.platform.startswith('win'):
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
    
    main()