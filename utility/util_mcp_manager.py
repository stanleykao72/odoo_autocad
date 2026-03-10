# -*- coding: utf-8 -*-
"""
MCP Manager — manages mcp_server_autocad.py lifecycle (in-process)

Replaces util_mcp_sse_manager.py. Runs the MCP server in a background thread
so it can share autocad_dispatcher and odoo_util instances with the GUI.
"""

import threading
import asyncio
import logging
from typing import Optional, Callable

logger = logging.getLogger(__name__)


class MCPManager:
    """Manages the MCP server in-process for GUI integration"""

    def __init__(self, autocad_dispatcher=None, odoo_util=None,
                 transport="sse", port=8084):
        self.autocad_dispatcher = autocad_dispatcher
        self.odoo_util = odoo_util
        self.transport = transport
        self.port = port
        self.is_running = False
        self._thread: Optional[threading.Thread] = None
        self._uvicorn_server = None
        self._status_callback: Optional[Callable] = None

        # Pass utility instances to mcp_server_autocad module (in-process)
        self._setup_shared_instances()

    def _setup_shared_instances(self):
        """Set shared instances on the mcp_server_autocad module"""
        try:
            import mcp_server_autocad
            if self.autocad_dispatcher:
                mcp_server_autocad.set_autocad_dispatcher(self.autocad_dispatcher)
            if self.odoo_util:
                mcp_server_autocad.set_odoo_util(self.odoo_util)
            logger.info("[MCP Manager] Shared instances configured")
        except Exception as e:
            logger.warning(f"[MCP Manager] Could not set shared instances: {e}")

    def set_status_callback(self, callback: Callable):
        """Set callback for status updates: callback(is_running: bool, message: str)"""
        self._status_callback = callback

    def _notify_status(self, is_running: bool, message: str):
        self.is_running = is_running
        if self._status_callback:
            try:
                self._status_callback(is_running, message)
            except Exception:
                pass

    def update_autocad_dispatcher(self, dispatcher):
        """Update the AutoCAD dispatcher (e.g. after mode switch)"""
        self.autocad_dispatcher = dispatcher
        try:
            import mcp_server_autocad
            mcp_server_autocad.set_autocad_dispatcher(dispatcher)
        except Exception as e:
            logger.warning(f"[MCP Manager] Could not update dispatcher: {e}")

    def start_server(self):
        """Start the MCP server in a background thread"""
        if self.is_running:
            logger.info("[MCP Manager] Server already running")
            return

        if self.transport == "sse":
            self._start_sse_in_thread()
        else:
            logger.info("[MCP Manager] stdio transport — no background server needed")

    def _start_sse_in_thread(self):
        """Start MCP SSE server in a background thread (in-process)"""
        def _run():
            try:
                import uvicorn
                import mcp_server_autocad

                # Get the Starlette app from our MCP server
                starlette_app = mcp_server_autocad.mcp.sse_app()

                config = uvicorn.Config(
                    starlette_app,
                    host="0.0.0.0",
                    port=self.port,
                    log_level="warning",
                )
                self._uvicorn_server = uvicorn.Server(config)

                logger.info(f"[MCP Manager] Starting MCP server on port {self.port}")

                # Create a new event loop for this thread
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)

                self._notify_status(True, f"MCP server running on port {self.port}")
                loop.run_until_complete(self._uvicorn_server.serve())
                loop.close()

            except Exception as e:
                import traceback
                err_detail = traceback.format_exc()
                logger.error(f"[MCP Manager] Server error: {e}\n{err_detail}")
                self._notify_status(False, f"MCP server error: {e}")
                return
            self._notify_status(False, "MCP server stopped")

        self._thread = threading.Thread(target=_run, daemon=True, name="mcp-server")
        self._thread.start()

    def stop_server(self):
        """Stop the MCP server"""
        if self._uvicorn_server:
            logger.info("[MCP Manager] Stopping MCP server...")
            self._uvicorn_server.should_exit = True
            self._uvicorn_server = None
        self._notify_status(False, "MCP server stopped")

    def restart_server(self):
        """Restart the MCP server"""
        self.stop_server()
        import time
        time.sleep(1)
        self.start_server()
