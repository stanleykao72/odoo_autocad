#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Integration test for FastMCP Server
"""

import sys
import os
import json
import subprocess
import time

# Add the project root to sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class FastMCPIntegrationTester:
    """Test the FastMCP server integration"""
    
    def __init__(self):
        self.server_path = os.path.join(os.path.dirname(__file__), '..', '..', 'mcp_server_fastmcp.py')
        self.server_process = None
    
    def start_server(self):
        """Start the FastMCP server process"""
        try:
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONDONTWRITEBYTECODE'] = '1'
            
            self.server_process = subprocess.Popen(
                [sys.executable, self.server_path],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=0,
                env=env
            )
            print(f"✓ FastMCP Server started with PID: {self.server_process.pid}")
            return True
        except Exception as e:
            print(f"✗ Failed to start server: {e}")
            return False
    
    def stop_server(self):
        """Stop the FastMCP server process"""
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
    
    def send_mcp_request(self, request_data, timeout=2):
        """Send MCP request to server"""
        try:
            request_json = json.dumps(request_data)
            print(f"→ Sending: {request_json}")
            
            self.server_process.stdin.write(request_json + '\n')
            self.server_process.stdin.flush()
            
            # Read response with timeout
            import select
            ready, _, _ = select.select([self.server_process.stdout], [], [], timeout)
            
            if ready:
                response_line = self.server_process.stdout.readline()
                if response_line:
                    response = json.loads(response_line.strip())
                    print(f"← Received: {json.dumps(response, indent=2)}")
                    return response
                else:
                    print("✗ Empty response received")
                    return None
            else:
                print(f"✗ No response within {timeout} seconds")
                return None
        except Exception as e:
            print(f"✗ Error sending request: {e}")
            return None
    
    def test_initialize_request(self):
        """Test MCP initialize request"""
        print("\n=== Testing FastMCP Initialize ===")
        request = {
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
                    }
                }
            }
        }
        
        response = self.send_mcp_request(request)
        if response and response.get('result', {}).get('protocolVersion'):
            print("✓ Initialize request successful")
            return True
        else:
            print("✗ Initialize request failed")
            return False
    
    def test_tools_list_request(self):
        """Test MCP tools/list request"""
        print("\n=== Testing FastMCP Tools List ===")
        request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list"
        }
        
        response = self.send_mcp_request(request)
        if response and 'result' in response and 'tools' in response['result']:
            tools = response['result']['tools']
            print(f"✓ Tools list successful - found {len(tools)} tools")
            for tool in tools:
                print(f"  - {tool.get('name', 'Unknown')}: {tool.get('description', 'No description')}")
            return True
        else:
            print("✗ Tools list failed")
            return False
    
    def test_tool_call_request(self):
        """Test MCP tools/call request"""
        print("\n=== Testing FastMCP Tool Call ===")
        request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "test_connection",
                "arguments": {}
            }
        }
        
        response = self.send_mcp_request(request)
        if response and 'result' in response:
            print("✓ Tool call successful")
            result = response['result']
            print(f"  Result: {result}")
            return True
        else:
            print("✗ Tool call failed")
            return False
    
    def run_integration_tests(self):
        """Run all integration tests"""
        print("🔄 Starting FastMCP Server Integration Tests...")
        
        if not self.start_server():
            return False
        
        try:
            # Give server time to start
            time.sleep(1)
            
            tests = [
                self.test_initialize_request,
                self.test_tools_list_request,
                self.test_tool_call_request
            ]
            
            passed = 0
            for test in tests:
                if test():
                    passed += 1
                time.sleep(0.5)  # Small delay between tests
            
            print(f"\n=== FastMCP Integration Test Results ===")
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
    
    tester = FastMCPIntegrationTester()
    success = tester.run_integration_tests()
    
    if success:
        print("\n🎉 All FastMCP integration tests passed!")
        print("   FastMCP Server is ready for use with Gemini CLI")
        sys.exit(0)
    else:
        print("\n❌ Some FastMCP integration tests failed!")
        sys.exit(1)