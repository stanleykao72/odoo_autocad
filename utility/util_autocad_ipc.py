# -*- coding: utf-8 -*-
"""
AutoCAD IPC 後端 — 透過 autocad-mcp File IPC 驅動 AutoCAD LT 2024+

提供與 UtilAutoCAD (COM) 相似的介面，但使用 File IPC 通訊。
"""

import asyncio
import json
import logging
import sys
import os

# Add autocad-mcp to path
_autocad_mcp_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "libs", "autocad-mcp", "src"
)
if _autocad_mcp_path not in sys.path:
    sys.path.insert(0, _autocad_mcp_path)

_logger = logging.getLogger(__name__)


class UtilAutoCADIPC:
    """AutoCAD IPC 後端 — 使用 autocad-mcp File IPC"""

    def __init__(self, log_util=None):
        self.log = log_util
        self._backend = None
        self._loop = None
        self.acad = None  # Compatibility: None means not COM-connected
        self.project_id = None
        self.project_name = None
        self.pr_no = None
        self.job_working_plan_id = None
        self.job_working_plan_name = None

    def _log(self, msg):
        if self.log:
            self.log.safe_log_insert(msg)
        else:
            _logger.info(msg.rstrip('\n'))

    def _get_loop(self):
        """Get or create an event loop for running async code"""
        if self._loop is None or self._loop.is_closed():
            self._loop = asyncio.new_event_loop()
        return self._loop

    def _run_async(self, coro):
        """Run an async coroutine synchronously"""
        loop = self._get_loop()
        return loop.run_until_complete(coro)

    async def _get_backend(self):
        """Lazy-init the File IPC backend"""
        if self._backend is None:
            try:
                from autocad_mcp.backends.file_ipc import FileIPCBackend
                self._backend = FileIPCBackend()
                await self._backend.initialize()
                self._log("[IPC] File IPC backend initialized\n")
            except Exception as e:
                self._log(f"[IPC] Failed to initialize File IPC backend: {e}\n")
                raise
        return self._backend

    async def _dispatch(self, command, params=None):
        """Send a command via File IPC and return the result"""
        backend = await self._get_backend()
        result = await backend._dispatch(command, params or {})
        return result

    # === Connection & Status ===

    def connected_autocad(self):
        """Check if IPC connection to AutoCAD is available"""
        try:
            result = self._run_async(self._dispatch("ping"))
            return result.ok if hasattr(result, 'ok') else bool(result)
        except Exception:
            return False

    def connect_autocad(self, main_body=None):
        """Establish IPC connection to AutoCAD"""
        self._log("[IPC] Connecting to AutoCAD via File IPC...\n")
        try:
            self._run_async(self._get_backend())
            # Test connection
            if self.connected_autocad():
                self._log("[IPC] AutoCAD IPC connection established\n")
            else:
                self._log("[IPC] AutoCAD not responding to IPC ping\n")
        except Exception as e:
            self._log(f"[IPC] Connection failed: {e}\n")

    # === Layout Management ===

    def get_active_layout(self):
        """Get the current active layout name"""
        try:
            result = self._run_async(self._dispatch("drawing-info"))
            if hasattr(result, 'ok') and result.ok and result.payload:
                return result.payload.get('active_layout')
            return None
        except Exception as e:
            self._log(f"[IPC] get_active_layout failed: {e}\n")
            return None

    def get_doc_layouts(self):
        """Get list of layout names (excluding Model)"""
        try:
            result = self._run_async(self._dispatch("drawing-info"))
            if hasattr(result, 'ok') and result.ok and result.payload:
                layouts = result.payload.get('layouts', [])
                return [l for l in layouts if l != 'Model']
            return []
        except Exception as e:
            self._log(f"[IPC] get_doc_layouts failed: {e}\n")
            return []

    # === Odoo-specific Operations (via ob_mcp_dispatch.lsp) ===

    def get_layouts_values(self):
        """Extract TABLE + Block data from all layouts (Odoo action)"""
        try:
            result = self._run_async(self._dispatch("odoo_extract_tables"))
            if hasattr(result, 'ok') and result.ok:
                return result.payload if result.payload else []
            self._log(f"[IPC] get_layouts_values failed: {getattr(result, 'error', 'unknown')}\n")
            return []
        except Exception as e:
            self._log(f"[IPC] get_layouts_values error: {e}\n")
            return []

    def get_layouts_header_id_to_pr(self):
        """Collect all header_ids from TABLEs (Odoo action)"""
        try:
            result = self._run_async(self._dispatch("odoo_get_header_ids"))
            if hasattr(result, 'ok') and result.ok:
                return result.payload if result.payload else {}
            return {}
        except Exception as e:
            self._log(f"[IPC] get_layouts_header_id_to_pr error: {e}\n")
            return {}

    def set_layouts_tables_id(self, boq_list):
        """Write header_id + detail_id back to TABLEs (Odoo action)"""
        try:
            # boq_list is the response from Odoo import2boq
            payload = {"all": boq_list} if isinstance(boq_list, list) else boq_list
            result = self._run_async(self._dispatch("odoo_write_ids", payload))
            if hasattr(result, 'ok') and result.ok:
                self._log("[IPC] IDs written back to TABLEs\n")
            else:
                self._log(f"[IPC] ID writeback failed: {getattr(result, 'error', 'unknown')}\n")
        except Exception as e:
            self._log(f"[IPC] set_layouts_tables_id error: {e}\n")

    def get_block_attributes(self):
        """Read attribute block values from current layout (Odoo action)"""
        try:
            result = self._run_async(self._dispatch("odoo_get_block_attrs"))
            if hasattr(result, 'ok') and result.ok:
                return result.payload if result.payload else {}
            return {}
        except Exception as e:
            self._log(f"[IPC] get_block_attributes error: {e}\n")
            return {}

    def set_block_attributes(self, attrs, layout_name=None):
        """Write attributes to Block (Odoo action)"""
        try:
            params = dict(attrs) if not isinstance(attrs, dict) else attrs
            if layout_name:
                params["layout_name"] = layout_name
            result = self._run_async(self._dispatch("odoo_set_block_attrs", params))
            if hasattr(result, 'ok') and result.ok:
                self._log("[IPC] Block attributes written\n")
            else:
                self._log(f"[IPC] set_block_attributes failed: {getattr(result, 'error', 'unknown')}\n")
        except Exception as e:
            self._log(f"[IPC] set_block_attributes error: {e}\n")

    def clear_table_id(self, layout=None):
        """Clear TABLE IDs"""
        try:
            result = self._run_async(self._dispatch("odoo_clear_ids"))
            if hasattr(result, 'ok') and result.ok:
                self._log("[IPC] TABLE IDs cleared\n")
        except Exception as e:
            self._log(f"[IPC] clear_table_id error: {e}\n")

    def clear_all_tables_id(self):
        """Clear TABLE IDs in all layouts"""
        self.clear_table_id()

    # === Drawing Operations (via autocad-mcp built-in commands) ===

    def draw_line(self, start_point, end_point, layer="0"):
        """Draw a line via IPC"""
        try:
            result = self._run_async(self._dispatch("create-line", {
                "x1": start_point[0], "y1": start_point[1],
                "x2": end_point[0], "y2": end_point[1],
                "layer": layer
            }))
            return hasattr(result, 'ok') and result.ok
        except Exception:
            return False

    def draw_circle(self, center_point, radius, layer="0"):
        """Draw a circle via IPC"""
        try:
            result = self._run_async(self._dispatch("create-circle", {
                "cx": center_point[0], "cy": center_point[1],
                "radius": radius, "layer": layer
            }))
            return hasattr(result, 'ok') and result.ok
        except Exception:
            return False

    def set_layer(self, layer_name, color=7, create_if_not_exist=True):
        """Set or create a layer via IPC"""
        try:
            if create_if_not_exist:
                self._run_async(self._dispatch("layer-create", {
                    "name": layer_name, "color": color
                }))
            self._run_async(self._dispatch("layer-set-current", {
                "name": layer_name
            }))
        except Exception as e:
            self._log(f"[IPC] set_layer error: {e}\n")

    def list_layers(self, filter_type="all", sort_by="name", include_details=True):
        """List layers via IPC"""
        try:
            result = self._run_async(self._dispatch("layer-list"))
            if hasattr(result, 'ok') and result.ok:
                return {"success": True, "layers": result.payload}
            return {"success": False, "layers": []}
        except Exception:
            return {"success": False, "layers": []}

    def scan_elements(self, element_type="all", **kwargs):
        """Scan drawing elements via IPC"""
        try:
            result = self._run_async(self._dispatch("entity-list", {
                "type": element_type
            }))
            if hasattr(result, 'ok') and result.ok:
                return {"success": True, "elements": result.payload}
            return {"success": False, "elements": []}
        except Exception:
            return {"success": False, "elements": []}
