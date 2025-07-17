#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Integration test for MCP SSE Server
"""

import sys
import os
import json
import requests
import threading
import time
import subprocess

# Add the project root to sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class SSEServerIntegrationTester:
    """Test the SSE server integration"""
    
    def __init__(self):
        self.server_process = None
        self.server_url = "http://localhost:8000"
        self.sse_endpoint = f"{self.server_url}/sse"
    
    def start_server(self, port=8080):
        """Start the SSE server in a separate process"""
        try:
            server_path = os.path.join(os.path.dirname(__file__), '..', '..', 'mcp_server_sse.py')
            self.server_url = f"http://localhost:{port}"
            self.sse_endpoint = f"{self.server_url}/sse"
            
            self.server_process = subprocess.Popen(
                [sys.executable, server_path, str(port)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # Wait for server to start
            time.sleep(2)
            
            # Check if server is running
            if self.server_process.poll() is None:
                print(f"✓ SSE Server started on port {port}")
                return True
            else:
                print("✗ SSE Server failed to start")
                stdout, stderr = self.server_process.communicate()
                print(f"STDOUT: {stdout}")
                print(f"STDERR: {stderr}")
                return False
                
        except Exception as e:
            print(f"✗ Failed to start SSE server: {e}")
            return False
    
    def stop_server(self):
        """Stop the SSE server"""
        if self.server_process:
            try:
                self.server_process.terminate()
                self.server_process.wait(timeout=5)
                print("✓ SSE Server stopped")
            except subprocess.TimeoutExpired:
                self.server_process.kill()
                print("✓ SSE Server killed")
            except Exception as e:
                print(f"✗ Error stopping server: {e}")
    
    def test_http_post_endpoint(self):
        """Test HTTP POST endpoint for MCP requests"""
        print("\n=== Testing HTTP POST Endpoint ===")
        
        try:
            # Test initialize request
            initialize_request = {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2024-11-05",
                    "clientInfo": {
                        "name": "Test Client",
                        "version": "1.0.0"
                    },
                    "capabilities": {
                        "roots": {
                            "listChanged": True
                        },
                        "sampling": {}
                    }
                }
            }
            
            response = requests.post(
                self.sse_endpoint,
                json=initialize_request,
                timeout=5
            )
            
            if response.status_code == 200:
                result = response.json()
                if result.get("result", {}).get("protocolVersion") == "2024-11-05":
                    print("✓ Initialize request successful")
                    return True
                else:
                    print(f"✗ Initialize request failed: {result}")
                    return False
            else:
                print(f"✗ HTTP POST failed with status {response.status_code}")
                return False
                
        except Exception as e:
            print(f"✗ Error testing HTTP POST: {e}")
            return False
    
    def test_tools_list_endpoint(self):
        """Test tools/list endpoint"""
        print("\n=== Testing Tools List Endpoint ===")
        
        try:
            tools_request = {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/list"
            }
            
            response = requests.post(
                self.sse_endpoint,
                json=tools_request,
                timeout=5
            )
            
            if response.status_code == 200:
                result = response.json()
                tools = result.get("result", {}).get("tools", [])
                if len(tools) > 0:
                    print(f"✓ Tools list successful - found {len(tools)} tools")
                    for tool in tools:
                        print(f"  - {tool.get('name', 'Unknown')}")
                    return True
                else:
                    print("✗ No tools found in response")
                    return False
            else:
                print(f"✗ Tools list failed with status {response.status_code}")
                return False
                
        except Exception as e:
            print(f"✗ Error testing tools list: {e}")
            return False
    
    def test_tool_call_endpoint(self):
        """Test tools/call endpoint"""
        print("\n=== Testing Tool Call Endpoint ===")
        
        try:
            tool_call_request = {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "tools/call",
                "params": {
                    "name": "test_connection",
                    "arguments": {}
                }
            }
            
            response = requests.post(
                self.sse_endpoint,
                json=tool_call_request,
                timeout=5
            )
            
            if response.status_code == 200:
                result = response.json()
                if "result" in result:
                    print("✓ Tool call successful")
                    print(f"  Result: {result['result']}")
                    return True
                else:
                    print(f"✗ Tool call failed: {result}")
                    return False
            else:
                print(f"✗ Tool call failed with status {response.status_code}")
                return False
                
        except Exception as e:
            print(f"✗ Error testing tool call: {e}")
            return False
    
    def test_sse_stream_endpoint(self):
        """Test SSE streaming endpoint"""
        print("\n=== Testing SSE Stream Endpoint ===")
        
        try:
            # Make a quick check to see if the SSE endpoint responds
            response = requests.get(self.sse_endpoint, timeout=2, stream=True)
            
            if response.status_code == 200:
                print("✓ SSE endpoint accessible")
                
                # Check for SSE headers
                content_type = response.headers.get("content-type", "")
                if "text/event-stream" in content_type or "text/plain" in content_type:
                    print("✓ SSE headers present")
                    return True
                else:
                    print(f"✗ Unexpected content type: {content_type}")
                    return False
            else:
                print(f"✗ SSE endpoint failed with status {response.status_code}")
                return False
                
        except Exception as e:
            print(f"✗ Error testing SSE stream: {e}")
            return False
    
    def run_integration_tests(self):
        """Run all integration tests"""
        print("🔄 Starting SSE Server Integration Tests...")
        
        if not self.start_server():
            return False
        
        try:
            tests = [
                self.test_http_post_endpoint,
                self.test_tools_list_endpoint,
                self.test_tool_call_endpoint,
                self.test_sse_stream_endpoint
            ]
            
            passed = 0
            for test in tests:
                if test():
                    passed += 1
                time.sleep(0.5)  # Small delay between tests
            
            print(f"\n=== SSE Integration Test Results ===")
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
    
    tester = SSEServerIntegrationTester()
    success = tester.run_integration_tests()
    
    if success:
        print("\n🎉 All SSE integration tests passed!")
        print("   SSE Server is ready for use")
        sys.exit(0)
    else:
        print("\n❌ Some SSE integration tests failed!")
        sys.exit(1)