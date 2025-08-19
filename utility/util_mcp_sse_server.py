# -*- coding: utf-8 -*-
"""
Standard MCP SSE Server Implementation
基於官方 MCP Python SDK 的標準 SSE 協定實作
"""

import asyncio
import json
import logging
import threading
import time
from typing import Any, Dict, List, Optional, Callable, Union
import sys
import os
from pathlib import Path
import traceback

# FastAPI and SSE dependencies
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# MCP Protocol types
from dataclasses import dataclass
from enum import Enum

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class MCPMessageType(str, Enum):
    """MCP Message Types according to specification"""
    REQUEST = "request"
    RESPONSE = "response" 
    NOTIFICATION = "notification"

@dataclass
class MCPMessage:
    """Standard MCP Message format"""
    jsonrpc: str = "2.0"
    id: Optional[Union[str, int]] = None
    method: Optional[str] = None
    params: Optional[Dict[str, Any]] = None
    result: Optional[Any] = None
    error: Optional[Dict[str, Any]] = None

class StandardMCPSSEServer:
    """
    Standard MCP SSE Server implementing official MCP protocol
    標準 MCP SSE 伺服器，實作官方 MCP 協定
    """
    
    def __init__(self, port: int = 8084, host: str = "localhost"):
        self.port = port
        self.host = host
        self.app: Optional[FastAPI] = None
        self.server: Optional[uvicorn.Server] = None
        self.server_thread: Optional[threading.Thread] = None
        self.is_running = False
        self.client_connections = set()
        self.tools = {}
        self.resources = {}
        self.prompts = {}
        
        # Set up logging
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        # Initialize MCP capabilities
        self.server_info = {
            "name": "AutoCAD-Odoo Integration MCP Server",
            "version": "5.1.0"
        }
        
        self.capabilities = {
            "logging": {},
            "prompts": {"listChanged": True},
            "resources": {"subscribe": True, "listChanged": True},
            "tools": {"listChanged": True}
        }
        
        # Initialize FastAPI app
        self._setup_fastapi_app()
        
        # Register default tools
        self._register_default_tools()
        
        self.logger.info(f"[MCP SSE] 標準 MCP SSE 伺服器初始化完成 (端口: {port})")
    
    def _setup_fastapi_app(self):
        """Setup FastAPI application with MCP endpoints"""
        self.app = FastAPI(
            title="AutoCAD-Odoo MCP SSE Server",
            description="Standard MCP Server with SSE transport for AutoCAD-Odoo integration",
            version="5.1.0"
        )
        
        # Add CORS middleware
        self.app.add_middleware(
            CORSMiddleware,
            allow_origins=["*"],
            allow_credentials=True,
            allow_methods=["*"],
            allow_headers=["*"],
        )
        
        # Root endpoint
        @self.app.get("/")
        async def root():
            return {
                "name": self.server_info["name"],
                "version": self.server_info["version"],
                "protocol": "MCP",
                "transport": ["sse", "stdio"],
                "capabilities": list(self.capabilities.keys())
            }
        
        # Health check
        @self.app.get("/health")
        async def health():
            return {
                "status": "healthy",
                "server": self.server_info,
                "capabilities": self.capabilities,
                "active_connections": len(self.client_connections)
            }
        
        # MCP JSON-RPC over HTTP endpoint  
        @self.app.post("/messages")
        async def handle_mcp_message(request: Request):
            try:
                # Get sessionId from query parameters (optional for backward compatibility)
                session_id = request.query_params.get("sessionId")
                
                # Only validate sessionId if it's provided and we have active SSE connections
                if session_id and len(self.client_connections) > 0 and session_id not in self.client_connections:
                    self.logger.warning(f"[MCP SSE] Session {session_id} not found, continuing without session validation")
                
                data = await request.json()
                self.logger.debug(f"[MCP SSE] 收到 MCP 訊息 (Session {session_id}): {data}")
                
                # Validate JSON-RPC format
                if not isinstance(data, dict) or data.get("jsonrpc") != "2.0":
                    raise HTTPException(400, "Invalid JSON-RPC format")
                
                # Handle the message
                response = await self._handle_mcp_request(data)
                self.logger.debug(f"[MCP SSE] 回應 MCP 訊息: {response}")
                
                return response
                
            except Exception as e:
                self.logger.error(f"[MCP SSE] 處理 MCP 訊息時發生錯誤: {e}")
                return JSONResponse(
                    status_code=400,
                    content={
                        "jsonrpc": "2.0",
                        "id": data.get("id") if 'data' in locals() else None,
                        "error": {
                            "code": -32600,
                            "message": "Invalid Request",
                            "data": str(e)
                        }
                    }
                )
        
        # SSE endpoint for real-time updates
        @self.app.get("/sse")
        async def sse_endpoint(request: Request):
            """SSE endpoint for MCP protocol over Server-Sent Events"""
            session_id = request.query_params.get("sessionId", str(time.time()))
            
            return StreamingResponse(
                self._sse_stream(request, session_id),
                media_type="text/event-stream",
                headers={
                    "Cache-Control": "no-cache",
                    "Connection": "keep-alive",
                    "Access-Control-Allow-Origin": "*",
                    "Access-Control-Allow-Headers": "*",
                }
            )
    
    async def _sse_stream(self, request: Request, session_id: str):
        """Generate SSE stream for MCP protocol"""
        self.client_connections.add(session_id)
        self.logger.info(f"[MCP SSE] 新的 SSE 連接: {session_id}")
        
        try:
            # Send initial connection established message
            init_event = {
                "jsonrpc": "2.0", 
                "method": "notifications/initialized",
                "params": {
                    "serverInfo": self.server_info,
                    "capabilities": self.capabilities,
                    "sessionId": session_id
                }
            }
            
            yield f"data: {json.dumps(init_event)}\n\n"
            
            # Keep connection alive and listen for disconnection
            while True:
                # Check if client is still connected
                if await request.is_disconnected():
                    break
                
                # Send keepalive every 30 seconds
                keepalive_event = {
                    "jsonrpc": "2.0",
                    "method": "notifications/ping", 
                    "params": {
                        "timestamp": time.time(),
                        "sessionId": session_id
                    }
                }
                
                yield f"data: {json.dumps(keepalive_event)}\n\n"
                await asyncio.sleep(30)
                
        except Exception as e:
            self.logger.error(f"[MCP SSE] SSE 串流錯誤 (Session {session_id}): {e}")
        finally:
            self.client_connections.discard(session_id)
            self.logger.info(f"[MCP SSE] SSE 連接已斷開: {session_id}")
    
    async def _handle_mcp_request(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Handle MCP JSON-RPC request"""
        method = data.get("method")
        params = data.get("params", {})
        request_id = data.get("id")
        
        try:
            if method == "initialize":
                return await self._handle_initialize(request_id, params)
            elif method == "tools/list":
                return await self._handle_tools_list(request_id, params)
            elif method == "tools/call":
                return await self._handle_tools_call(request_id, params)
            elif method == "resources/list":
                return await self._handle_resources_list(request_id, params)
            elif method == "prompts/list":
                return await self._handle_prompts_list(request_id, params)
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Method not found: {method}"
                    }
                }
        except Exception as e:
            self.logger.error(f"[MCP SSE] 處理方法 {method} 時發生錯誤: {e}")
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": "Internal error",
                    "data": str(e)
                }
            }
    
    async def _handle_initialize(self, request_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle MCP initialize method"""
        client_info = params.get("clientInfo", {})
        protocol_version = params.get("protocolVersion")
        client_capabilities = params.get("capabilities", {})
        
        self.logger.info(f"[MCP SSE] 初始化客戶端: {client_info.get('name', 'Unknown')} v{client_info.get('version', 'Unknown')}")
        self.logger.info(f"[MCP SSE] 協定版本: {protocol_version}")
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": self.capabilities,
                "serverInfo": self.server_info
            }
        }
    
    async def _handle_tools_list(self, request_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle tools/list method"""
        tools_list = [
            {
                "name": name,
                "description": tool_info["description"],
                "inputSchema": tool_info["schema"]
            }
            for name, tool_info in self.tools.items()
        ]
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "tools": tools_list
            }
        }
    
    async def _handle_tools_call(self, request_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle tools/call method"""
        tool_name = params.get("name")
        arguments = params.get("arguments", {})
        
        if tool_name not in self.tools:
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32602,
                    "message": f"Tool not found: {tool_name}"
                }
            }
        
        try:
            tool_func = self.tools[tool_name]["function"]
            result = await self._call_tool_function(tool_func, arguments)
            
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps(result, ensure_ascii=False, indent=2)
                        }
                    ]
                }
            }
            
        except Exception as e:
            self.logger.error(f"[MCP SSE] 執行工具 {tool_name} 時發生錯誤: {e}")
            return {
                "jsonrpc": "2.0", 
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": f"Tool execution failed: {str(e)}"
                }
            }
    
    async def _handle_resources_list(self, request_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle resources/list method"""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "resources": list(self.resources.values())
            }
        }
    
    async def _handle_prompts_list(self, request_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle prompts/list method"""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "prompts": list(self.prompts.values())
            }
        }
    
    async def _call_tool_function(self, tool_func: Callable, arguments: Dict[str, Any]) -> Any:
        """Call tool function, handling both sync and async functions"""
        if asyncio.iscoroutinefunction(tool_func):
            return await tool_func(**arguments)
        else:
            return tool_func(**arguments)
    
    def register_tool(self, name: str, description: str, schema: Dict[str, Any], function: Callable):
        """Register a tool with the MCP server"""
        self.tools[name] = {
            "description": description,
            "schema": schema,
            "function": function
        }
        self.logger.info(f"[MCP SSE] 註冊工具: {name}")
    
    def _register_default_tools(self):
        """Register default tools for testing and basic functionality"""
        
        # Test connection tool
        def test_connection():
            return {
                "status": "connected",
                "server": self.server_info["name"],
                "version": self.server_info["version"],
                "timestamp": time.time(),
                "message": "MCP 連接測試成功"
            }
        
        self.register_tool(
            name="test_connection",
            description="測試 MCP 伺服器連接狀態",
            schema={
                "type": "object",
                "properties": {},
                "required": []
            },
            function=test_connection
        )
        
        # Get server info tool
        def get_server_info():
            return {
                "server_info": self.server_info,
                "capabilities": self.capabilities,
                "active_connections": len(self.client_connections),
                "available_tools": list(self.tools.keys()),
                "uptime": time.time()
            }
        
        self.register_tool(
            name="get_server_info",
            description="獲取伺服器資訊和功能清單",
            schema={
                "type": "object", 
                "properties": {},
                "required": []
            },
            function=get_server_info
        )
        
        # Natural language command processor
        def process_natural_language_command(command: str, user_context: Dict[str, Any] = None):
            """Process natural language drawing commands"""
            try:
                # Import NLP processor
                from utility.util_nlp_processor import NaturalLanguageProcessor
                
                processor = NaturalLanguageProcessor()
                intent = processor.parse_drawing_intent(command)
                
                # Format response for MCP
                return {
                    "command": command,
                    "parsed_intent": {
                        "action": intent.action,
                        "parameters": intent.parameters,
                        "confidence": intent.confidence
                    },
                    "suggestions": intent.suggestions,
                    "processing_time": "< 0.5秒",
                    "status": "解析成功" if intent.confidence > 0.7 else "需要更多資訊"
                }
                
            except Exception as e:
                self.logger.error(f"[MCP SSE] 自然語言處理錯誤: {e}")
                return {
                    "command": command,
                    "error": str(e),
                    "status": "處理失敗"
                }
        
        self.register_tool(
            name="process_natural_language_command",
            description="處理自然語言繪圖指令 (支援中文)",
            schema={
                "type": "object",
                "properties": {
                    "command": {
                        "type": "string",
                        "description": "自然語言繪圖指令，如 '畫一個半徑10的圓形'"
                    },
                    "user_context": {
                        "type": "object",
                        "description": "用戶上下文資訊 (可選)",
                        "default": {}
                    }
                },
                "required": ["command"]
            },
            function=process_natural_language_command
        )
        
        # AutoCAD status check
        def check_autocad_status():
            """Check AutoCAD connection status"""
            try:
                # Try to import AutoCAD utility
                from utility.util_autocad import UtilAutoCAD
                from utility.util_log import UtilLog
                
                # Create a temporary log util for testing (None parent for console mode)
                log_util = UtilLog(None)
                autocad_util = UtilAutoCAD(None, log_util)
                
                # Test connection by checking if AutoCAD is connected
                is_connected = autocad_util.connected_autocad()
                
                if is_connected and autocad_util.acad:
                    status = {
                        "connected": True,
                        "application": str(autocad_util.acad.Name) if hasattr(autocad_util.acad, 'Name') else "AutoCAD",
                        "version": str(autocad_util.acad.Version) if hasattr(autocad_util.acad, 'Version') else "Unknown",
                        "active_document": str(autocad_util.doc.Name) if autocad_util.doc and hasattr(autocad_util.doc, 'Name') else "None"
                    }
                else:
                    status = {
                        "connected": False,
                        "application": "AutoCAD",
                        "version": "Unknown", 
                        "active_document": "None"
                    }
                
                return {
                    "autocad_connected": status.get("connected", False),
                    "application_name": status.get("application", "Unknown"),
                    "version": status.get("version", "Unknown"),
                    "active_document": status.get("active_document", "None"),
                    "status": "已連接" if status.get("connected") else "未連接"
                }
                
            except Exception as e:
                return {
                    "autocad_connected": False,
                    "error": str(e),
                    "status": "連接失敗"
                }
        
        self.register_tool(
            name="check_autocad_status", 
            description="檢查 AutoCAD 連接狀態",
            schema={
                "type": "object",
                "properties": {},
                "required": []
            },
            function=check_autocad_status
        )
    
    def start_server(self) -> bool:
        """Start the MCP SSE server"""
        if self.is_running:
            self.logger.warning("[MCP SSE] 伺服器已在運行中")
            return True
        
        try:
            self.logger.info(f"[MCP SSE] 啟動標準 MCP SSE 伺服器 (端口: {self.port})")
            
            # Create uvicorn server
            config = uvicorn.Config(
                app=self.app,
                host=self.host,
                port=self.port,
                log_level="warning",  # Reduce log noise
                access_log=False,
                log_config=None
            )
            
            self.server = uvicorn.Server(config)
            
            # Start server in separate thread
            self.server_thread = threading.Thread(
                target=self._run_server_thread,
                daemon=True,
                name=f"MCP-SSE-{self.port}"
            )
            
            self.server_thread.start()
            
            # Wait for server to start
            time.sleep(2)
            
            # Verify server is running
            if self._check_server_health():
                self.is_running = True
                self.logger.info(f"[MCP SSE] ✅ 標準 MCP SSE 伺服器啟動成功 (端口: {self.port})")
                return True
            else:
                self.logger.error("[MCP SSE] ❌ 伺服器啟動後健康檢查失敗")
                return False
                
        except Exception as e:
            self.logger.error(f"[MCP SSE] ❌ 啟動伺服器時發生錯誤: {e}")
            return False
    
    def stop_server(self) -> bool:
        """Stop the MCP SSE server"""
        if not self.is_running:
            return True
            
        try:
            self.logger.info("[MCP SSE] 停止標準 MCP SSE 伺服器")
            
            # Stop uvicorn server
            if self.server:
                self.server.should_exit = True
            
            # Wait for thread to finish
            if self.server_thread and self.server_thread.is_alive():
                self.server_thread.join(timeout=5)
            
            # Clean up
            self.is_running = False
            self.client_connections.clear()
            self.server = None
            self.server_thread = None
            
            self.logger.info("[MCP SSE] ✅ 伺服器已停止")
            return True
            
        except Exception as e:
            self.logger.error(f"[MCP SSE] ❌ 停止伺服器時發生錯誤: {e}")
            return False
    
    def _run_server_thread(self):
        """Run uvicorn server in thread with proper async setup"""
        try:
            # Set up event loop for this thread
            if sys.platform == "win32":
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            
            # Run server
            loop.run_until_complete(self.server.serve())
            
        except Exception as e:
            self.logger.error(f"[MCP SSE] 伺服器執行緒錯誤: {e}")
        finally:
            self.logger.info("[MCP SSE] 伺服器執行緒結束")
    
    def _check_server_health(self) -> bool:
        """Check if server is responding to requests"""
        try:
            import requests
            
            url = f"http://{self.host}:{self.port}/health"
            response = requests.get(url, timeout=5)
            
            if response.status_code == 200:
                data = response.json()
                return data.get("status") == "healthy"
                
        except Exception as e:
            self.logger.debug(f"[MCP SSE] 健康檢查失敗: {e}")
            
        return False
    
    def get_server_status(self) -> Dict[str, Any]:
        """Get current server status"""
        return {
            "is_running": self.is_running,
            "host": self.host,
            "port": self.port,
            "server_info": self.server_info,
            "capabilities": self.capabilities,
            "active_connections": len(self.client_connections),
            "registered_tools": list(self.tools.keys()),
            "health_check": self._check_server_health() if self.is_running else False
        }
    
    def test_mcp_connection(self) -> Dict[str, Any]:
        """Test MCP connection to the server"""
        if not self.is_running:
            return {"success": False, "error": "伺服器未運行"}
            
        try:
            import requests
            
            # Test the standard MCP endpoints
            url = f"http://{self.host}:{self.port}/messages"
            
            # Test initialize
            init_request = {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize", 
                "params": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {},
                    "clientInfo": {"name": "test-client", "version": "1.0.0"}
                }
            }
            
            response = requests.post(url, json=init_request, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if "result" in data:
                    # Test tools/list
                    tools_request = {
                        "jsonrpc": "2.0",
                        "id": 2,
                        "method": "tools/list",
                        "params": {}
                    }
                    
                    tools_response = requests.post(url, json=tools_request, timeout=10)
                    if tools_response.status_code == 200:
                        tools_data = tools_response.json()
                        if "result" in tools_data:
                            tools = tools_data["result"]["tools"]
                            return {
                                "success": True,
                                "tools_count": len(tools),
                                "tools": [tool["name"] for tool in tools],
                                "protocol": "Standard MCP JSON-RPC"
                            }
            
            return {
                "success": False,
                "error": f"HTTP {response.status_code}: {response.text[:100]}"
            }
            
        except Exception as e:
            return {"success": False, "error": str(e)}

def main():
    """Test the MCP SSE server"""
    import sys
    import signal
    
    # Create and start server
    server = StandardMCPSSEServer(port=8084)
    
    def signal_handler(sig, frame):
        print("\n🛑 正在停止 MCP SSE 伺服器...")
        server.stop_server()
        sys.exit(0)
    
    signal.signal(signal.SIGINT, signal_handler)
    
    if server.start_server():
        print(f"🚀 標準 MCP SSE 伺服器已啟動於 http://localhost:8084")
        print("📋 可用端點:")
        print("  • http://localhost:8084/ - 伺服器資訊")
        print("  • http://localhost:8084/health - 健康檢查")
        print("  • http://localhost:8084/messages - MCP JSON-RPC 端點")
        print("  • http://localhost:8084/sse - SSE 串流")
        print("\n按 Ctrl+C 停止伺服器")
        
        try:
            # Keep main thread alive
            while server.is_running:
                time.sleep(1)
        except KeyboardInterrupt:
            pass
    else:
        print("❌ 伺服器啟動失敗")
        sys.exit(1)

if __name__ == "__main__":
    main()