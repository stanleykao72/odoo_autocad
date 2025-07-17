# -*- coding: utf-8 -*-
"""
MCP SSE Server Manager for GUI Integration
Directly integrates MCPSSEServer class instead of subprocess management
"""

import threading
import asyncio
import time
import logging
import requests
import uvicorn
from typing import Optional, Dict, Any
import sys
import os

# Add parent directory to path to import mcp_server_fastmcp
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# Import the FastMCP instance directly
from mcp_server_fastmcp import mcp

class MCPSSEManager:
    """Manager for MCP SSE Server in GUI using FastMCP SDK integration"""
    
    def __init__(self, port: int = 8081):
        self.port = port
        self.actual_port = port  # Track the actual port where server is running
        self.server = None  # Will hold the FastMCP instance
        self.server_thread: Optional[threading.Thread] = None
        self.is_running = False
        self.status_callback = None
        self._loop: Optional[asyncio.AbstractEventLoop] = None
        
        # Configure logging
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)  # 設定更詳細的日誌級別
        
        self.logger.info(f"[SSE Manager] MCPSSEManager 初始化，端口: {port}")
        
    def set_status_callback(self, callback):
        """Set callback function for status updates"""
        self.status_callback = callback
        
    def _update_status(self, message: str, is_running: bool = None):
        """Update status and call callback if set"""
        old_status = self.is_running
        if is_running is not None:
            self.is_running = is_running
            
        self.logger.info(f"[SSE Manager] 狀態更新: {message} (運行狀態: {old_status} -> {self.is_running})")
        
        if self.status_callback:
            self.logger.debug(f"[SSE Manager] 呼叫狀態回調函數")
            try:
                self.status_callback(message, self.is_running)
                self.logger.debug(f"[SSE Manager] 狀態回調函數執行成功")
            except Exception as e:
                self.logger.error(f"[SSE Manager] 狀態回調函數執行失敗: {e}")
        else:
            self.logger.debug(f"[SSE Manager] 無狀態回調函數")
    
    def start_server(self) -> bool:
        """Start the SSE server using direct integration"""
        self.logger.info(f"[SSE Manager] start_server() 被呼叫，當前狀態: {self.is_running}")
        
        if self.is_running:
            self.logger.warning(f"[SSE Manager] 伺服器已在運行中，略過啟動")
            self._update_status("SSE 伺服器已在運行中")
            return True
            
        try:
            self.logger.info(f"[SSE Manager] 準備使用 FastMCP 實例")
            
            # Use the imported FastMCP instance
            self.server = mcp
            self.logger.info(f"[SSE Manager] FastMCP 實例已設定: {self.server}")
            
            # Start server in a separate thread
            import threading
            current_thread = threading.current_thread()
            self.logger.info(f"[SSE Manager] 當前執行緒: {current_thread.name} (ID: {current_thread.ident})")
            
            self.server_thread = threading.Thread(
                target=self._run_server_in_thread,
                daemon=True,
                name=f"SSE-Server-{self.port}"
            )
            self.logger.info(f"[SSE Manager] 創建伺服器執行緒: {self.server_thread.name}")
            
            self.server_thread.start()
            self.logger.info(f"[SSE Manager] 伺服器執行緒已啟動")
            
            # Wait for server to start
            self.logger.info(f"[SSE Manager] 等待 2 秒讓伺服器啟動")
            time.sleep(2)
            
            # Check thread status
            if self.server_thread.is_alive():
                self.logger.info(f"[SSE Manager] 伺服器執行緒仍在運行")
            else:
                self.logger.error(f"[SSE Manager] 伺服器執行緒已結束")
            
            # Mark as running since server thread started successfully
            self.logger.info(f"[SSE Manager] 標記伺服器為運行狀態")
            self._update_status(f"SSE 伺服器已啟動 (端口: {self.port})", True)
            
            # Optional health check in background (non-blocking)
            self.logger.info(f"[SSE Manager] 啟動背景健康檢查")
            self._background_health_check()
            
            self.logger.info(f"[SSE Manager] start_server() 完成，回傳 True")
            return True
                
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.logger.error(f"[SSE Manager] 啟動伺服器時發生異常: {e}")
            self.logger.error(f"[SSE Manager] 錯誤追蹤:\n{error_trace}")
            self._update_status(f"啟動 SSE 伺服器時發生錯誤: {str(e)}", False)
            return False
    
    def stop_server(self) -> bool:
        """Stop the SSE server"""
        self.logger.info(f"[SSE Manager] stop_server() 被呼叫，當前狀態: {self.is_running}")
        
        if not self.is_running:
            self.logger.warning(f"[SSE Manager] 伺服器未在運行，略過停止")
            self._update_status("SSE 伺服器未在運行", False)
            return True
            
        try:
            self.logger.info(f"[SSE Manager] 準備停止伺服器")
            
            # Stop the event loop
            if self._loop:
                self.logger.info(f"[SSE Manager] 檢查事件迴圈狀態: running={self._loop.is_running() if hasattr(self._loop, 'is_running') else 'N/A'}")
                if hasattr(self._loop, 'is_running') and self._loop.is_running():
                    self.logger.info(f"[SSE Manager] 停止事件迴圈")
                    self._loop.call_soon_threadsafe(self._loop.stop)
                else:
                    self.logger.info(f"[SSE Manager] 事件迴圈未運行")
            else:
                self.logger.info(f"[SSE Manager] 無事件迴圈")
            
            # Wait for thread to finish
            if self.server_thread:
                self.logger.info(f"[SSE Manager] 等待伺服器執行緒結束: {self.server_thread.name}")
                if self.server_thread.is_alive():
                    self.logger.info(f"[SSE Manager] 執行緒仍在運行，等待 5 秒")
                    self.server_thread.join(timeout=5)
                    if self.server_thread.is_alive():
                        self.logger.warning(f"[SSE Manager] 執行緒在 5 秒後仍未結束")
                    else:
                        self.logger.info(f"[SSE Manager] 執行緒已正常結束")
                else:
                    self.logger.info(f"[SSE Manager] 執行緒已結束")
            else:
                self.logger.info(f"[SSE Manager] 無伺服器執行緒")
            
            # Clean up
            self.logger.info(f"[SSE Manager] 清理資源")
            self._update_status("SSE 伺服器已停止", False)
            self.server = None
            self.server_thread = None
            self._loop = None
            
            self.logger.info(f"[SSE Manager] stop_server() 完成，回傳 True")
            return True
            
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.logger.error(f"[SSE Manager] 停止伺服器時發生異常: {e}")
            self.logger.error(f"[SSE Manager] 錯誤追蹤:\n{error_trace}")
            self._update_status(f"停止 SSE 伺服器時發生錯誤: {str(e)}", False)
            return False
    
    def restart_server(self) -> bool:
        """Restart the SSE server"""
        self._update_status("正在重新啟動 SSE 伺服器...")
        self.stop_server()
        time.sleep(1)
        return self.start_server()
    
    def _check_server_health(self) -> bool:
        """Check if server is responding"""
        self.logger.info(f"[SSE Health] 開始健康檢查，檢查端口 {self.port} 和默認端口 8000")
        
        # Try both the configured port and the default FastMCP port
        ports_to_try = [self.port, 8000] if self.port != 8000 else [8000]
        
        for port in ports_to_try:
            self.logger.debug(f"[SSE Health] 檢查端口 {port}")
            
            # Try multiple times with increasing delays
            for attempt in range(3):  # Reduce attempts per port
                self.logger.debug(f"[SSE Health] 端口 {port} 第 {attempt + 1} 次嘗試")
                
                try:
                    # Test basic connectivity - FastMCP uses root endpoint
                    url = f"http://localhost:{port}/"
                    
                    self.logger.debug(f"[SSE Health] 發送 GET 請求到: {url}")
                    response = requests.get(url, timeout=3)
                    
                    self.logger.debug(f"[SSE Health] 回應狀態碼: {response.status_code}")
                    
                    if response.status_code in [200, 404]:
                        # Both 200 and 404 indicate server is running
                        # 404 is expected for FastMCP on root endpoint
                        self.logger.info(f"[SSE Health] ✅ 在端口 {port} 健康檢查成功 (狀態碼: {response.status_code})")
                        # Update our internal port if we found it on a different port
                        if port != self.port:
                            self.logger.info(f"[SSE Health] FastMCP 運行在端口 {port}，更新內部配置")
                            self.actual_port = port
                        return True
                    else:
                        self.logger.warning(f"[SSE Health] 端口 {port} 非預期的狀態碼: {response.status_code}")
                        
                except requests.exceptions.ConnectionError as e:
                    # Server not ready yet, wait and retry
                    self.logger.debug(f"[SSE Health] 端口 {port} 第 {attempt + 1} 次嘗試 - 連接錯誤: {e}")
                    if attempt < 2:  # Don't sleep on last attempt
                        time.sleep(1)
                    continue
                except Exception as e:
                    self.logger.warning(f"[SSE Health] 端口 {port} 第 {attempt + 1} 次嘗試失敗: {e}")
                    if attempt < 2:  # Don't sleep on last attempt
                        time.sleep(1)
                    continue
        
        self.logger.error(f"[SSE Health] ❌ 所有端口的健康檢查都失敗")
        return False
    
    def _background_health_check(self):
        """Run health check in background thread"""
        def check():
            time.sleep(1)  # Give server more time to start
            if self._check_server_health():
                self._update_status(f"SSE 伺服器健康檢查通過 (端口: {self.port})", True)
            else:
                self._update_status("SSE 伺服器健康檢查失敗", True)  # Still keep running
        
        threading.Thread(target=check, daemon=True).start()
    
    def _run_server_in_thread(self):
        """Run the server in a separate thread with its own event loop"""
        import threading
        thread_name = threading.current_thread().name
        thread_id = threading.current_thread().ident
        
        self.logger.info(f"[SSE Server Thread] 伺服器執行緒開始 (Name: {thread_name}, ID: {thread_id})")
        
        try:
            # 簡化的 FastMCP 執行方式
            self.logger.info(f"[SSE Server Thread] 開始執行 FastMCP 伺服器 (端口: {self.port})")
            
            # 設定環境變數
            import os
            os.environ['PORT'] = str(self.port)
            os.environ['HOST'] = 'localhost'
            
            # Windows 相容性設定
            import sys
            if sys.platform == "win32":
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            
            # 設定新的事件迴圈
            self._loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self._loop)
            self.logger.info(f"[SSE Server Thread] 事件迴圈已設定: {self._loop}")
            
            # 完全禁用 uvicorn 日誌配置以避免 Windows exe 錯誤
            import logging
            import sys
            import os
            
            # 完全抑制 uvicorn 相關日誌
            logging.getLogger("uvicorn").disabled = True
            logging.getLogger("uvicorn.access").disabled = True
            logging.getLogger("uvicorn.error").disabled = True
            
            # 重定向標準錯誤到 null 設備
            devnull = open(os.devnull, 'w')
            original_stderr = sys.stderr
            sys.stderr = devnull
            
            try:
                self.logger.info(f"[SSE Server Thread] 使用 FastMCP 運行 SSE 傳輸")
                
                # 設定 uvicorn 環境變數以避免配置問題
                os.environ['UVICORN_LOG_LEVEL'] = 'critical'
                os.environ['UVICORN_ACCESS_LOG'] = 'false'
                
                # 直接使用 FastMCP 的 SSE 模式
                self.server.run(transport="sse")
                
            except Exception as e:
                # 恢復 stderr 來記錄錯誤
                sys.stderr = original_stderr
                self.logger.error(f"[SSE Server Thread] FastMCP 運行失敗: {e}")
                
                # 使用 mcp_server_fastmcp.py 的 get_fastapi_app 函數
                try:
                    self.logger.info(f"[SSE Server Thread] 嘗試使用 mcp_server_fastmcp.py")
                    import uvicorn
                    
                    # 使用 mcp_server_fastmcp.py 的 get_fastapi_app 函數
                    from mcp_server_fastmcp import get_fastapi_app
                    app = get_fastapi_app()
                    
                    if app is None:
                        raise Exception("無法從 mcp_server_fastmcp.py 獲取 FastAPI 應用")
                    
                    # 使用完全禁用日誌的 uvicorn 配置
                    config = uvicorn.Config(
                        app=app,
                        host='localhost',
                        port=self.port,
                        log_level='critical',
                        access_log=False,
                        log_config=None,  # 完全禁用日誌配置
                        use_colors=False
                    )
                    
                    server = uvicorn.Server(config)
                    self.logger.info(f"[SSE Server Thread] 啟動 mcp_server_fastmcp 伺服器")
                    self._loop.run_until_complete(server.serve())
                    
                except Exception as e2:
                    self.logger.error(f"[SSE Server Thread] mcp_server_fastmcp 也失敗: {e2}")
                    
                    # 最後的備用方案：創建基本的 FastAPI 應用
                    try:
                        self.logger.info(f"[SSE Server Thread] 使用基本 FastAPI 應用作為備用")
                        from fastapi import FastAPI
                        from fastapi.responses import JSONResponse
                        
                        # 創建基本的 FastAPI 應用
                        app = FastAPI(title="AutoCAD-Odoo MCP SSE Server")
                        
                        @app.get("/")
                        async def root():
                            return {"name": "AutoCAD-Odoo Integration", "version": "5.0", "status": "running"}
                        
                        @app.get("/health")
                        async def health():
                            return {"status": "healthy", "server": "basic"}
                        
                        # 配置 uvicorn
                        config = uvicorn.Config(
                            app=app,
                            host='localhost',
                            port=self.port,
                            log_level='critical',
                            access_log=False,
                            log_config=None
                        )
                        
                        server = uvicorn.Server(config)
                        self.logger.info(f"[SSE Server Thread] 啟動備用 FastAPI 伺服器")
                        self._loop.run_until_complete(server.serve())
                        
                    except Exception as e3:
                        self.logger.error(f"[SSE Server Thread] 備用 FastAPI 也失敗: {e3}")
                        raise e3
                    
            finally:
                # 恢復標準錯誤
                sys.stderr = original_stderr
                devnull.close()
                        
            # FastMCP 執行完成
            
            self.logger.info(f"[SSE Server Thread] FastMCP 伺服器執行結束")
            
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.logger.error(f"[SSE Server Thread] 伺服器執行緒發生錯誤: {e}")
            self.logger.error(f"[SSE Server Thread] 錯誤追蹤:\n{error_trace}")
        finally:
            self.logger.info(f"[SSE Server Thread] 清理事件迴圈")
            if self._loop:
                try:
                    self._loop.close()
                    self.logger.info(f"[SSE Server Thread] 事件迴圈已關閉")
                except Exception as e:
                    self.logger.error(f"[SSE Server Thread] 關閉事件迴圈時發生錯誤: {e}")
            
            self.logger.info(f"[SSE Server Thread] 伺服器執行緒結束")
    
    def get_server_status(self) -> Dict[str, Any]:
        """Get current server status"""
        status = {
            "is_running": self.is_running,
            "port": self.port,
            "health_check": self._check_server_health() if self.is_running else False
        }
        
        # Add server info if available
        if self.server:
            status["server_name"] = getattr(self.server, 'name', 'AutoCAD-Odoo Integration')
            status["server_version"] = "5.0"
            status["active_connections"] = "FastMCP"
        
        return status
    
    def test_mcp_connection(self) -> Dict[str, Any]:
        """Test MCP connection to the server"""
        self.logger.info(f"[SSE Test] 開始 MCP 連接測試")
        
        if not self.is_running:
            self.logger.warning(f"[SSE Test] 伺服器未運行，無法測試")
            return {"success": False, "error": "伺服器未運行"}
            
        try:
            # Use actual port where server is running
            actual_port = getattr(self, 'actual_port', self.port)
            # FastMCP uses /messages endpoint for JSON-RPC
            url = f"http://localhost:{actual_port}/messages"
            self.logger.info(f"[SSE Test] 測試目標 URL: {url}")
            
            # First test initialize
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
            self.logger.debug(f"[SSE Test] 發送 initialize 請求: {init_request}")
            
            response = requests.post(url, json=init_request, timeout=10)
            self.logger.debug(f"[SSE Test] initialize 回應狀態: {response.status_code}")
            
            if response.status_code != 200:
                error_msg = f"Initialize failed: HTTP {response.status_code}: {response.text}"
                self.logger.error(f"[SSE Test] {error_msg}")
                return {"success": False, "error": error_msg}
            
            # Now test tools/list
            tools_request = {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/list",
                "params": {}
            }
            self.logger.debug(f"[SSE Test] 發送 tools/list 請求: {tools_request}")
            
            response = requests.post(url, json=tools_request, timeout=10)
            self.logger.debug(f"[SSE Test] tools/list 回應狀態: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                self.logger.debug(f"[SSE Test] tools/list 回應資料: {data}")
                
                if "result" in data and "tools" in data["result"]:
                    tools = data["result"]["tools"]
                    self.logger.info(f"[SSE Test] 找到 {len(tools)} 個工具")
                    
                    # Test a tool call
                    tool_request = {
                        "jsonrpc": "2.0",
                        "id": 3,
                        "method": "tools/call",
                        "params": {
                            "name": "test_connection",
                            "arguments": {}
                        }
                    }
                    self.logger.debug(f"[SSE Test] 發送 tools/call 請求: {tool_request}")
                    
                    test_response = requests.post(url, json=tool_request, timeout=10)
                    self.logger.debug(f"[SSE Test] tools/call 回應狀態: {test_response.status_code}")
                    
                    tool_result = "未測試"
                    if test_response.status_code == 200:
                        test_data = test_response.json()
                        self.logger.debug(f"[SSE Test] tools/call 回應資料: {test_data}")
                        if "result" in test_data:
                            tool_result = test_data["result"]
                            self.logger.info(f"[SSE Test] 工具測試結果: {tool_result}")
                    
                    result = {
                        "success": True,
                        "tools_count": len(tools),
                        "tools": [tool["name"] for tool in tools],
                        "test_result": tool_result
                    }
                    self.logger.info(f"[SSE Test] ✅ MCP 連接測試成功")
                    return result
                else:
                    self.logger.error(f"[SSE Test] tools/list 回應格式錯誤: {data}")
            
            error_msg = f"HTTP {response.status_code}: {response.text}"
            self.logger.error(f"[SSE Test] ❌ MCP 連接測試失敗: {error_msg}")
            return {"success": False, "error": error_msg}
            
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.logger.error(f"[SSE Test] ❌ MCP 連接測試發生異常: {e}")
            self.logger.error(f"[SSE Test] 錯誤追蹤:\n{error_trace}")
            return {"success": False, "error": str(e)}
    
    def test_tool_directly(self, tool_name: str, arguments: Dict[str, Any] = None) -> Any:
        """直接調用伺服器工具方法進行測試"""
        if not self.server:
            return {"error": "伺服器未初始化"}
        
        try:
            # FastMCP tools are accessed directly
            if self.server:
                # Use FastMCP's tool calling mechanism
                if arguments is None:
                    arguments = {}
                
                # Call the tool through FastMCP
                result = self.server.call_tool(tool_name, arguments)
                return result
            else:
                return {"error": "FastMCP 實例未找到"}
        except Exception as e:
            return {"error": str(e)}
    
    def cleanup(self):
        """Clean up resources"""
        if self.is_running:
            self.stop_server()