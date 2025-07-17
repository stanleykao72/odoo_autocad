#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test MCP Initialize Protocol
Tests the MCP initialization and tools listing protocol
"""

import json
import requests
import time

def test_mcp_initialize():
    """Test MCP initialization protocol"""
    print("Testing MCP initialization protocol...")
    
    # Test initialize request
    initialize_request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "clientInfo": {
                "name": "test-client",
                "version": "1.0.0"
            },
            "capabilities": {}
        }
    }
    
    print("Sending initialize request...")
    try:
        response = requests.post(
            "http://localhost:8083/mcp",
            headers={"Content-Type": "application/json"},
            json=initialize_request,
            timeout=10
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Initialize response: {json.dumps(result, indent=2)}")
            
            # Check if tools capability is properly set
            capabilities = result.get("result", {}).get("capabilities", {})
            tools_capability = capabilities.get("tools", {})
            
            if tools_capability:
                print(f"✅ Tools capability found: {tools_capability}")
            else:
                print("❌ Tools capability is empty!")
                
        else:
            print(f"❌ Initialize failed: {response.status_code}")
            print(f"Response: {response.text}")
            
    except Exception as e:
        print(f"❌ Initialize request failed: {e}")
    
    # Test tools/list request
    tools_request = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list"
    }
    
    print("\nSending tools/list request...")
    try:
        response = requests.post(
            "http://localhost:8083/mcp",
            headers={"Content-Type": "application/json"},
            json=tools_request,
            timeout=10
        )
        
        if response.status_code == 200:
            result = response.json()
            tools = result.get("result", {}).get("tools", [])
            print(f"✅ Tools list response:")
            print(f"   Found {len(tools)} tools:")
            for tool in tools:
                print(f"   - {tool.get('name')}: {tool.get('description')}")
        else:
            print(f"❌ Tools list failed: {response.status_code}")
            print(f"Response: {response.text}")
            
    except Exception as e:
        print(f"❌ Tools list request failed: {e}")

if __name__ == "__main__":
    test_mcp_initialize()