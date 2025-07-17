#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test MCP Connection for Gemini CLI
Tests the MCP server connection and protocol handling
"""

import sys
import json
import subprocess
import threading
import time
import pytest
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))


class TestMCPConnection:
    """Test MCP server connection functionality"""
    
    def test_mcp_server_starts_with_enable_mcp(self):
        """Test that MCP server can start with --enable-mcp flag"""
        # Start the application with --enable-mcp
        cmd = [sys.executable, "odoo.py", "--enable-mcp"]
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=0
        )
        
        # Give it time to start
        time.sleep(2)
        
        # Check if process is still running
        assert process.poll() is None, "Process should still be running"
        
        # Clean up
        process.terminate()
        process.wait(timeout=5)
    
    def test_mcp_stdio_protocol(self):
        """Test basic MCP stdio protocol communication"""
        # Simple test server that responds to MCP requests
        test_script = '''
import sys
import json

while True:
    try:
        line = sys.stdin.readline()
        if not line:
            break
        
        request = json.loads(line.strip())
        
        if request.get('method') == 'initialize':
            response = {
                "jsonrpc": "2.0",
                "id": request.get('id'),
                "result": {
                    "protocolVersion": "1.0",
                    "capabilities": {
                        "tools": {}
                    }
                }
            }
            print(json.dumps(response))
            sys.stdout.flush()
        elif request.get('method') == 'tools/list':
            response = {
                "jsonrpc": "2.0",
                "id": request.get('id'),
                "result": []
            }
            print(json.dumps(response))
            sys.stdout.flush()
            
    except Exception as e:
        pass
'''
        
        # Create a test process
        process = subprocess.Popen(
            [sys.executable, "-c", test_script],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Send initialize request
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {}
        }
        process.stdin.write(json.dumps(init_request) + "\n")
        process.stdin.flush()
        
        # Read response
        response_line = process.stdout.readline()
        response = json.loads(response_line)
        
        assert response["id"] == 1
        assert "result" in response
        assert "protocolVersion" in response["result"]
        
        # Clean up
        process.terminate()
        process.wait()


def create_test_mcp_wrapper():
    """Create a wrapper script for testing MCP with proper stdio handling"""
    wrapper_content = '''#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""MCP Wrapper for Testing"""

import sys
import os
import subprocess
import threading

def forward_stdio():
    """Forward stdio between parent and child process"""
    # Set unbuffered mode
    os.environ['PYTHONUNBUFFERED'] = '1'
    
    # Start the actual application
    cmd = [sys.executable, "odoo.py", "--enable-mcp"]
    
    process = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,  # Suppress stderr to avoid protocol interference
        text=True,
        bufsize=0
    )
    
    # Forward stdin to process
    def stdin_to_process():
        try:
            while True:
                line = sys.stdin.readline()
                if not line:
                    break
                process.stdin.write(line)
                process.stdin.flush()
        except:
            pass
    
    # Forward process stdout to our stdout
    def process_to_stdout():
        try:
            while True:
                line = process.stdout.readline()
                if not line:
                    break
                sys.stdout.write(line)
                sys.stdout.flush()
        except:
            pass
    
    # Start forwarding threads
    stdin_thread = threading.Thread(target=stdin_to_process, daemon=True)
    stdout_thread = threading.Thread(target=process_to_stdout, daemon=True)
    
    stdin_thread.start()
    stdout_thread.start()
    
    # Wait for process
    try:
        process.wait()
    except KeyboardInterrupt:
        process.terminate()

if __name__ == "__main__":
    forward_stdio()
'''
    
    wrapper_path = Path(__file__).parent.parent.parent / "mcp_wrapper.py"
    with open(wrapper_path, "w", encoding="utf-8") as f:
        f.write(wrapper_content)
    
    return wrapper_path


if __name__ == "__main__":
    # Run tests
    pytest.main([__file__, "-v"])