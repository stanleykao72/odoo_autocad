#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script for improved MCP server
"""

import sys
import os
import json
import subprocess
import time
import threading
from io import StringIO

# Add the project root to sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class MCPServerTester:
    """Test the improved MCP server"""
    
    def __init__(self):
        self.server_path = os.path.join(os.path.dirname(__file__), '..', '..', 'mcp_server_improved.py')
        self.server_process = None
        self.server_output = StringIO()
        self.server_error = StringIO()
    
    def start_server(self):
        """Start the MCP server process"""
        try:
            self.server_process = subprocess.Popen(
                [sys.executable, self.server_path],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=0
            )
            print(f"✓ Server started with PID: {self.server_process.pid}")
            return True
        except Exception as e:
            print(f"✗ Failed to start server: {e}")
            return False
    
    def stop_server(self):
        """Stop the MCP server process"""
        if self.server_process:
            try:
                self.server_process.terminate()
                self.server_process.wait(timeout=5)
                print("✓ Server stopped")
            except subprocess.TimeoutExpired:
                self.server_process.kill()
                print("✓ Server killed")
            except Exception as e:
                print(f"✗ Error stopping server: {e}")
    
    def send_request(self, request_data):
        """Send JSON-RPC request to server"""
        try:
            request_json = json.dumps(request_data)
            print(f"→ Sending: {request_json}")
            
            self.server_process.stdin.write(request_json + '\n')
            self.server_process.stdin.flush()
            
            # Read response with timeout
            response_line = self.server_process.stdout.readline()
            if response_line:
                response = json.loads(response_line.strip())
                print(f"← Received: {json.dumps(response, indent=2)}")
                return response
            else:
                print("✗ No response received")
                return None
        except Exception as e:
            print(f"✗ Error sending request: {e}")
            return None
    
    def test_initialize(self):
        """Test initialize request"""
        print("\n=== Testing Initialize ===")
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "clientInfo": {
                    "name": "test-client",
                    "version": "1.0.0"
                }
            }
        }
        
        response = self.send_request(request)
        if response and response.get('result', {}).get('protocolVersion') == "2024-11-05":
            print("✓ Initialize test passed")
            return True
        else:
            print("✗ Initialize test failed")
            return False
    
    def test_tools_list(self):
        """Test tools/list request"""
        print("\n=== Testing Tools List ===")
        request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list"
        }
        
        response = self.send_request(request)
        if response and 'result' in response and 'tools' in response['result']:
            tools = response['result']['tools']
            print(f"✓ Tools list test passed - found {len(tools)} tools")
            for tool in tools:
                print(f"  - {tool['name']}: {tool['description']}")
            return True
        else:
            print("✗ Tools list test failed")
            return False
    
    def test_tools_call(self):
        """Test tools/call request"""
        print("\n=== Testing Tools Call ===")
        request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "test_connection",
                "arguments": {}
            }
        }
        
        response = self.send_request(request)
        if response and 'result' in response and 'content' in response['result']:
            content = response['result']['content']
            print(f"✓ Tools call test passed")
            for item in content:
                print(f"  Content: {item.get('text', '')}")
            return True
        else:
            print("✗ Tools call test failed")
            return False
    
    def test_ping(self):
        """Test ping request"""
        print("\n=== Testing Ping ===")
        request = {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "ping"
        }
        
        response = self.send_request(request)
        if response and 'result' in response and response['result'].get('status') == 'ok':
            print("✓ Ping test passed")
            return True
        else:
            print("✗ Ping test failed")
            return False
    
    def test_error_handling(self):
        """Test error handling"""
        print("\n=== Testing Error Handling ===")
        request = {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "nonexistent_method"
        }
        
        response = self.send_request(request)
        if response and 'error' in response and response['error']['code'] == -32601:
            print("✓ Error handling test passed")
            return True
        else:
            print("✗ Error handling test failed")
            return False
    
    def run_all_tests(self):
        """Run all tests"""
        print("Starting MCP Server Tests...")
        
        if not self.start_server():
            return False
        
        try:
            # Give server time to start
            time.sleep(0.5)
            
            tests = [
                self.test_initialize,
                self.test_tools_list,
                self.test_tools_call,
                self.test_ping,
                self.test_error_handling
            ]
            
            passed = 0
            for test in tests:
                if test():
                    passed += 1
                time.sleep(0.1)  # Small delay between tests
            
            print(f"\n=== Test Results ===")
            print(f"Passed: {passed}/{len(tests)}")
            print(f"Success rate: {passed/len(tests)*100:.1f}%")
            
            return passed == len(tests)
        
        finally:
            self.stop_server()


if __name__ == "__main__":
    # Set UTF-8 encoding for Windows console
    if sys.platform.startswith('win'):
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
    
    tester = MCPServerTester()
    success = tester.run_all_tests()
    
    if success:
        print("\n🎉 All tests passed!")
        sys.exit(0)
    else:
        print("\n❌ Some tests failed!")
        sys.exit(1)