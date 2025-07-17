#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test Gemini CLI SSE Behavior
Tests to understand how Gemini CLI interacts with SSE server
"""

import asyncio
import json
import httpx
import time
from httpx_sse import connect_sse

async def test_gemini_behavior():
    """Test SSE connection like Gemini CLI"""
    print("Testing SSE connection behavior like Gemini CLI...")
    
    # Create HTTP client
    client = httpx.AsyncClient(timeout=30.0)
    
    try:
        # Step 1: Establish SSE connection
        print("\n1. Establishing SSE connection...")
        async with connect_sse(client, "GET", "http://localhost:8083/sse") as sse:
            print("   SSE connected, waiting for events...")
            
            # Listen for initial events
            event_count = 0
            async for event in sse.aiter_sse():
                event_count += 1
                print(f"   Event #{event_count}: {event.event} = {event.data}")
                
                # Stop after a few events to continue with test
                if event_count >= 3:
                    break
        
        # Step 2: Send initialize request via POST
        print("\n2. Sending initialize request via POST...")
        initialize_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "clientInfo": {
                    "name": "gemini-cli",
                    "version": "0.1.0"
                },
                "capabilities": {}
            }
        }
        
        response = await client.post(
            "http://localhost:8083/sse",
            json=initialize_request,
            headers={"Content-Type": "application/json"}
        )
        
        print(f"   Response status: {response.status_code}")
        print(f"   Response body: {response.text}")
        
        # Step 3: Send tools/list request
        print("\n3. Sending tools/list request...")
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list"
        }
        
        response = await client.post(
            "http://localhost:8083/sse",
            json=tools_request,
            headers={"Content-Type": "application/json"}
        )
        
        print(f"   Response status: {response.status_code}")
        print(f"   Response body: {response.text}")
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        await client.aclose()

async def test_sse_with_messages():
    """Test SSE bidirectional communication"""
    print("\nTesting SSE bidirectional communication...")
    
    async with httpx.AsyncClient(timeout=30.0) as client:
        # Maybe Gemini expects messages through SSE events?
        print("Checking if server sends MCP messages through SSE...")
        
        async with connect_sse(client, "GET", "http://localhost:8083/sse") as sse:
            print("Listening for SSE events for 5 seconds...")
            
            start_time = time.time()
            async for event in sse.aiter_sse():
                print(f"Event: {event.event}")
                print(f"Data: {event.data}")
                print(f"ID: {event.id}")
                print("-" * 40)
                
                # Check if this looks like an MCP message
                try:
                    data = json.loads(event.data)
                    if "jsonrpc" in data:
                        print("*** Found JSON-RPC message in SSE! ***")
                except:
                    pass
                
                # Stop after 5 seconds
                if time.time() - start_time > 5:
                    break

if __name__ == "__main__":
    print("=" * 60)
    print("Gemini CLI SSE Behavior Test")
    print("=" * 60)
    
    try:
        # First install required package
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "httpx", "httpx-sse"], check=True)
    except:
        pass
    
    # Run tests
    asyncio.run(test_gemini_behavior())
    asyncio.run(test_sse_with_messages())