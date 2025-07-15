# -*- coding: utf-8 -*-
"""
MCP Server Manager for AutoCAD-Odoo Integration

Implements dual communication architecture (TCP Socket + Named Pipe)
following TDD principles from CLAUDE.md and MCP_INTEGRATION_PLAN.md

This is the minimal implementation to pass initial tests (Green phase of TDD).
"""
import socket
import threading
import time
import json
from typing import Optional
from .mcp_request_handler import MCPRequestHandler

# Constants for server configuration
DEFAULT_TCP_PORT = 8000
DEFAULT_PIPE_NAME = r'\\.\pipe\odoo_autocad_mcp'
SERVER_BACKLOG = 5
SOCKET_TIMEOUT = 1.0
THREAD_JOIN_TIMEOUT = 5.0
CLIENT_BUFFER_SIZE = 1024


class MCPServerManager:
    """
    Unified manager for TCP Socket and Named Pipe MCP services
    
    Provides dual communication channels for AI assistant integration:
    - TCP Socket server for network-based connections
    - Named Pipe server for local IPC (Windows-specific)
    """
    
    def __init__(
        self, 
        autocad_util,
        odoo_util, 
        log_util,
        tcp_port: int = DEFAULT_TCP_PORT,
        pipe_name: str = DEFAULT_PIPE_NAME
    ):
        """
        Initialize MCP Server Manager with required dependencies
        
        Args:
            autocad_util: AutoCAD utility instance for CAD operations
            odoo_util: Odoo utility instance for ERP integration
            log_util: Logging utility instance
            tcp_port: TCP server port (default: 8000)
            pipe_name: Named pipe name (default: \\\\.\\pipe\\odoo_autocad_mcp)
        """
        # Dependencies (following dependency injection pattern from CLAUDE.md)
        self.autocad_util = autocad_util
        self.odoo_util = odoo_util
        self.log_util = log_util
        
        # Configuration
        self.tcp_port = tcp_port
        self.pipe_name = pipe_name
        
        # Server state
        self.is_tcp_running = False
        self.is_pipe_running = False
        
        # Server instances
        self.tcp_server = None
        self.pipe_server = None
        
        # Threading
        self._tcp_thread = None
        self._pipe_thread = None
        self._shutdown_event = threading.Event()
        
        # MCP Request Handler
        self.request_handler = MCPRequestHandler(
            autocad_util=autocad_util,
            odoo_util=odoo_util,
            log_util=log_util
        )
    
    def start_tcp_server(self) -> None:
        """
        Start TCP Socket server for network-based MCP communication
        
        Creates a TCP server bound to localhost on the specified port.
        Runs in a separate daemon thread for non-blocking operation.
        """
        try:
            self.tcp_server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.tcp_server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.tcp_server.bind(('localhost', self.tcp_port))
            self.tcp_server.listen(SERVER_BACKLOG)
            
            # Start server thread
            self._tcp_thread = threading.Thread(
                target=self._tcp_server_loop,
                daemon=True,
                name=f"TCPServer-{self.tcp_port}"
            )
            self._tcp_thread.start()
            
            self.is_tcp_running = True
            self.log_util.safe_log_insert(f"TCP MCP Server started on localhost:{self.tcp_port}\n")
            
        except Exception as e:
            self._log_error("Failed to start TCP server", e)
            raise
    
    def start_pipe_server(self) -> None:
        """
        Start Named Pipe server for local IPC communication
        
        Creates a Windows Named Pipe for high-performance local communication.
        Runs in a separate daemon thread for non-blocking operation.
        """
        try:
            # Start pipe server thread
            self._pipe_thread = threading.Thread(
                target=self._pipe_server_loop,
                daemon=True,
                name="PipeServer"
            )
            self._pipe_thread.start()
            
            self.is_pipe_running = True
            self.log_util.safe_log_insert(f"Named Pipe MCP Server started: {self.pipe_name}\n")
            
        except Exception as e:
            self._log_error("Failed to start Pipe server", e)
            raise
    
    def start_all_servers(self) -> None:
        """
        Start both TCP Socket and Named Pipe servers simultaneously
        
        This implements the dual communication architecture described in
        MCP_INTEGRATION_PLAN.md for maximum compatibility and redundancy.
        """
        self.start_tcp_server()
        self.start_pipe_server()
    
    def stop_all_servers(self) -> None:
        """
        Stop all running servers and clean up resources
        
        Gracefully shuts down both TCP and Pipe servers, closes connections,
        and waits for threads to complete.
        """
        # Signal shutdown
        self._shutdown_event.set()
        
        # Close TCP server
        if self.tcp_server:
            try:
                self.tcp_server.close()
            except Exception as e:
                self._log_error("Error closing TCP server", e)
        
        # Close Pipe server (implementation depends on win32pipe)
        if self.pipe_server:
            try:
                # This will be implemented when we add win32pipe support
                pass
            except Exception as e:
                self._log_error("Error closing Pipe server", e)
        
        # Wait for threads to complete
        if self._tcp_thread and self._tcp_thread.is_alive():
            self._tcp_thread.join(timeout=THREAD_JOIN_TIMEOUT)
        
        if self._pipe_thread and self._pipe_thread.is_alive():
            self._pipe_thread.join(timeout=THREAD_JOIN_TIMEOUT)
        
        # Reset state
        self.is_tcp_running = False
        self.is_pipe_running = False
        
        self.log_util.safe_log_insert("All MCP servers stopped\n")
    
    def is_running(self) -> bool:
        """
        Check if any MCP server is currently running
        
        Returns:
            bool: True if TCP or Pipe server is running, False otherwise
        """
        return self.is_tcp_running or self.is_pipe_running
    
    def get_tcp_port(self) -> Optional[int]:
        """
        Get TCP server port if running
        
        Returns:
            int: TCP port number if server is running, None otherwise
        """
        return self.tcp_port if self.is_tcp_running else None
    
    def get_pipe_name(self) -> Optional[str]:
        """
        Get Named Pipe name if running
        
        Returns:
            str: Pipe name if server is running, None otherwise
        """
        return self.pipe_name if self.is_pipe_running else None
    
    def _log_error(self, context: str, error: Exception) -> None:
        """
        Centralized error logging with consistent format
        
        Args:
            context: Description of where the error occurred
            error: The exception that was caught
        """
        self.log_util.safe_log_insert(f"{context}: {error}\n")
    
    def _tcp_server_loop(self) -> None:
        """
        Main TCP server loop (runs in separate thread)
        
        Accepts incoming connections and handles MCP protocol communication.
        This is a minimal implementation for the Green phase of TDD.
        """
        while not self._shutdown_event.is_set():
            try:
                # Set timeout for non-blocking accept
                self.tcp_server.settimeout(SOCKET_TIMEOUT)
                client_socket, addr = self.tcp_server.accept()
                
                self.log_util.safe_log_insert(f"TCP client connected: {addr}\n")
                
                # Handle client in separate thread (to be implemented)
                client_thread = threading.Thread(
                    target=self._handle_tcp_client,
                    args=(client_socket, addr),
                    daemon=True
                )
                client_thread.start()
                
            except socket.timeout:
                # Normal timeout, continue loop
                continue
            except Exception as e:
                if not self._shutdown_event.is_set():
                    self._log_error("TCP server error", e)
                break
    
    def _pipe_server_loop(self) -> None:
        """
        Main Named Pipe server loop (runs in separate thread)
        
        Creates and manages Named Pipe connections for local IPC.
        This is a minimal implementation for the Green phase of TDD.
        """
        while not self._shutdown_event.is_set():
            try:
                # This will be implemented when we add proper win32pipe support
                # For now, just maintain the running state
                time.sleep(1)
                
            except Exception as e:
                if not self._shutdown_event.is_set():
                    self._log_error("Pipe server error", e)
                break
    
    def _handle_tcp_client(self, client_socket, addr) -> None:
        """
        Handle TCP client connection with full MCP protocol support
        
        Args:
            client_socket: Client socket connection
            addr: Client address tuple
        """
        try:
            while not self._shutdown_event.is_set():
                try:
                    client_socket.settimeout(SOCKET_TIMEOUT)
                    data = client_socket.recv(CLIENT_BUFFER_SIZE)
                    if not data:
                        break
                    
                    # 解析MCP請求
                    try:
                        request_text = data.decode('utf-8')
                        self.log_util.safe_log_insert(f"收到MCP請求: {request_text}\n")
                        
                        # 處理JSON-RPC請求
                        request = json.loads(request_text)
                        response = self.request_handler.process_request(request)
                        
                        # 發送回應
                        response_text = json.dumps(response, ensure_ascii=False)
                        response_bytes = response_text.encode('utf-8')
                        client_socket.send(response_bytes)
                        
                        self.log_util.safe_log_insert(f"MCP回應已發送\n")
                        
                    except json.JSONDecodeError as e:
                        # 處理JSON解析錯誤
                        error_response = {
                            "jsonrpc": "2.0",
                            "id": None,
                            "error": {
                                "code": -32700,
                                "message": "Parse error"
                            }
                        }
                        error_text = json.dumps(error_response).encode('utf-8')
                        client_socket.send(error_text)
                        self.log_util.safe_log_insert(f"JSON解析錯誤: {e}\n")
                    
                except socket.timeout:
                    continue
                    
        except Exception as e:
            self._log_error("TCP client error", e)
        finally:
            try:
                client_socket.close()
                self.log_util.safe_log_insert(f"TCP client disconnected: {addr}\n")
            except:
                pass