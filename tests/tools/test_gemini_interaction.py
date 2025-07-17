#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script to simulate Gemini CLI interaction with MCP server
"""

import sys
import os
import json
import subprocess
import threading
import time
from queue import Queue, Empty

# Add the project root to sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class GeminiMCPTester:
    """Simulate Gemini CLI interaction with MCP server"""
    
    def __init__(self):
        self.server_path = os.path.join(os.path.dirname(__file__), '..', '..', 'mcp_server_improved.py')
        self.server_process = None
    
    def start_server_with_config(self):
        """Start server with the exact configuration from Gemini CLI"""
        try:
            # Use the exact same command that Gemini CLI would use
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONDONTWRITEBYTECODE'] = '1'
            
            self.server_process = subprocess.Popen(
                ['python', self.server_path],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=0,
                env=env
            )
            
            print(f"✓ Server started with PID: {self.server_process.pid}")
            print(f"✓ Environment: PYTHONUNBUFFERED=1, PYTHONDONTWRITEBYTECODE=1")
            return True
        except Exception as e:
            print(f"✗ Failed to start server: {e}")
            return False
    
    def stop_server(self):
        """Stop the server"""
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
    
    def simulate_gemini_handshake(self):
        """Simulate the handshake that Gemini CLI would do"""
        print("\n=== Simulating Gemini CLI Handshake ===")
        
        # Step 1: Initialize
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "clientInfo": {
                    "name": "Gemini CLI",
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
        
        response = self.send_request(init_request)
        if not response or 'error' in response:
            print("✗ Initialize failed")
            return False
        
        print("✓ Initialize successful")
        
        # Step 2: Get tools list
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list"
        }
        
        response = self.send_request(tools_request)
        if not response or 'error' in response:
            print("✗ Tools list failed")
            return False
        
        tools = response.get('result', {}).get('tools', [])
        print(f"✓ Tools list successful - {len(tools)} tools available")
        
        # Step 3: Test a tool call
        tool_request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "test_connection",
                "arguments": {}
            }
        }
        
        response = self.send_request(tool_request)
        if not response or 'error' in response:
            print("✗ Tool call failed")
            return False
        
        print("✓ Tool call successful")
        return True
    
    def send_request(self, request):
        """Send request and get response"""
        try:
            request_json = json.dumps(request)
            print(f"→ {request['method']}: {request_json}")
            
            self.server_process.stdin.write(request_json + '\n')
            self.server_process.stdin.flush()
            
            # Read response with timeout
            response_line = self.server_process.stdout.readline()
            if response_line:
                response = json.loads(response_line.strip())
                print(f"← Response: {json.dumps(response, indent=2)}")
                return response
            else:
                print("✗ No response received")
                return None
                
        except Exception as e:
            print(f"✗ Error in communication: {e}")
            return None
    
    def run_gemini_simulation(self):
        """Run the full Gemini CLI simulation"""
        print("🔄 Starting Gemini CLI Simulation...")
        
        if not self.start_server_with_config():
            return False
        
        try:
            # Give server time to start
            time.sleep(0.5)
            
            success = self.simulate_gemini_handshake()
            
            if success:
                print("\n✅ Gemini CLI simulation completed successfully!")
                print("   The server should now work with Gemini CLI")
            else:
                print("\n❌ Gemini CLI simulation failed!")
                print("   There may be compatibility issues")
            
            return success
            
        finally:
            self.stop_server()


if __name__ == "__main__":
    # Set UTF-8 encoding for Windows console
    if sys.platform.startswith('win'):
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
    
    tester = GeminiMCPTester()
    success = tester.run_gemini_simulation()
    
    if success:
        print("\n🎉 Ready for Gemini CLI!")
        print("   Try running: gemini chat")
        print("   The autocad-odoo server should now show as 'Connected'")
        sys.exit(0)
    else:
        print("\n❌ Needs troubleshooting")
        sys.exit(1)