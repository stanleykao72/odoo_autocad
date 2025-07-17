#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Simple test for FastMCP Server
"""

import sys
import os
import json
import subprocess
import time
import threading
from queue import Queue, Empty

# Add the project root to sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))


class SimpleFastMCPTester:
    """Simple test for FastMCP server"""
    
    def __init__(self):
        self.server_path = os.path.join(os.path.dirname(__file__), '..', '..', 'mcp_server_fastmcp.py')
        self.server_process = None
    
    def test_server_startup(self):
        """Test that the server can start up correctly"""
        print("🔄 Testing FastMCP Server Startup...")
        
        try:
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONDONTWRITEBYTECODE'] = '1'
            
            # Start server
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
            
            # Give it time to start
            time.sleep(2)
            
            # Check if it's still running
            if self.server_process.poll() is None:
                print("✓ Server is running")
                
                # Test with a simple request with proper capabilities
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
                            },
                            "sampling": {}
                        }
                    }
                }
                
                request_json = json.dumps(request)
                print(f"→ Sending: {request_json}")
                
                self.server_process.stdin.write(request_json + '\n')
                self.server_process.stdin.flush()
                
                # Wait a bit for response
                time.sleep(1)
                
                print("✓ Request sent successfully")
                return True
            else:
                print("✗ Server exited unexpectedly")
                stderr = self.server_process.stderr.read()
                if stderr:
                    print(f"Error: {stderr}")
                return False
                
        except Exception as e:
            print(f"✗ Error testing server: {e}")
            return False
        finally:
            if self.server_process:
                self.server_process.terminate()
                print("✓ Server terminated")
    
    def test_server_with_gemini_config(self):
        """Test server with the exact Gemini CLI configuration"""
        print("\n🔄 Testing with Gemini CLI Configuration...")
        
        try:
            # Use exact same command as in Gemini CLI config
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
            
            print(f"✓ Server started with Gemini CLI config")
            
            # Wait for startup
            time.sleep(2)
            
            if self.server_process.poll() is None:
                print("✓ Server is running with Gemini CLI config")
                return True
            else:
                print("✗ Server failed with Gemini CLI config")
                stderr = self.server_process.stderr.read()
                if stderr:
                    print(f"Error: {stderr}")
                return False
                
        except Exception as e:
            print(f"✗ Error with Gemini CLI config: {e}")
            return False
        finally:
            if self.server_process:
                self.server_process.terminate()
                print("✓ Server terminated")


if __name__ == "__main__":
    # Set UTF-8 encoding for Windows console
    if sys.platform.startswith('win'):
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
    
    tester = SimpleFastMCPTester()
    
    test1 = tester.test_server_startup()
    test2 = tester.test_server_with_gemini_config()
    
    if test1 and test2:
        print("\n🎉 All FastMCP tests passed!")
        print("   FastMCP Server is ready for Gemini CLI")
        sys.exit(0)
    else:
        print("\n❌ Some FastMCP tests failed!")
        sys.exit(1)