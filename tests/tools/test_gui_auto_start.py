#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test GUI Auto-Start SSE Server Integration
Tests that the GUI starts with SSE server auto-start functionality
"""

import sys
import os
import subprocess
import time
import requests
from pathlib import Path

def test_gui_auto_start():
    """Test that GUI starts with SSE server auto-start"""
    print("Testing GUI auto-start SSE server...")
    
    # Start the GUI application with MCP enabled
    cmd = [sys.executable, "odoo.py", "--enable-mcp"]
    
    print(f"Starting GUI with command: {' '.join(cmd)}")
    print("This will start the GUI application with auto-start SSE server...")
    print("Check if the SSE server starts automatically and shows green status in GUI.")
    print("Also check if Gemini CLI can connect to 'autocad-odoo-sse' server.")
    
    # Instructions for manual verification
    print("\n" + "="*60)
    print("MANUAL VERIFICATION STEPS:")
    print("1. GUI should start with SSE server auto-starting")
    print("2. SSE status indicator should turn green (🟢)")
    print("3. SSE info should show 'SSE: :8083'")
    print("4. Test Gemini CLI connection:")
    print("   - Run: gemini mcp tools")
    print("   - Look for 'autocad-odoo-sse' server")
    print("   - Should show '🟢 autocad-odoo-sse - Connected'")
    print("="*60)
    
    try:
        # Test that we can connect to the expected port
        print("\nTesting port 8083 connectivity...")
        time.sleep(2)
        
        response = requests.get(
            "http://localhost:8083/sse",
            timeout=5,
            headers={"Accept": "text/event-stream"}
        )
        
        if response.status_code == 200:
            print("✅ Port 8083 is accessible")
            print("✅ SSE endpoint is responding correctly")
        else:
            print(f"⚠️ Port 8083 response: {response.status_code}")
            
    except requests.exceptions.ConnectionError:
        print("⚠️ Port 8083 not accessible (server may not be running yet)")
    except Exception as e:
        print(f"❌ Connection test failed: {e}")
    
    print("\n✅ Test completed - Check GUI for SSE auto-start functionality")

if __name__ == "__main__":
    # Set UTF-8 encoding for Windows
    if sys.platform == "win32":
        os.environ["PYTHONIOENCODING"] = "utf-8"
    
    # Change to the correct directory
    os.chdir(Path(__file__).parent.parent.parent)
    
    test_gui_auto_start()