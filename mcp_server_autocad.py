# -*- coding: utf-8 -*-
"""
MCP Server for AutoCAD-Odoo Integration — 13 tools (8 AutoCAD + 5 Odoo)

Single server combining autocad-mcp's 8 upstream tools with 5 Odoo tools.
Supports stdio and streamable-http transports.

Tools (8 from autocad-mcp):
  drawing, entity, layer, block, annotation, pid, view, system

Tools (5 Odoo-specific):
  odoo_push_boq, odoo_create_pr, odoo_get_setup, odoo_get_colors, odoo_status
"""

import os
import sys
import logging
import builtins
from datetime import datetime
from typing import Dict, Any

# Add autocad-mcp to path
_autocad_mcp_src = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "libs", "autocad-mcp", "src"
)
if _autocad_mcp_src not in sys.path:
    sys.path.insert(0, _autocad_mcp_src)

# Inject ToolResult into builtins BEFORE importing autocad_mcp.server —
# upstream tools use `-> ToolResult` (= str | list) as return type annotation.
# Without this, FastMCP raises ForwardRef('ToolResult') error.
builtins.ToolResult = str | list

logger = logging.getLogger(__name__)

# Import upstream autocad-mcp server — this registers 8 tools on its mcp instance
from autocad_mcp.server import mcp  # noqa: E402

# Reuse the SAME FastMCP instance — our 5 Odoo tools are added below,
# giving a single server with all 13 tools.

# --- Global utility instances ---
_autocad_dispatcher = None
_odoo_util = None


def set_autocad_dispatcher(dispatcher):
    """Set the AutoCAD dispatcher (UtilAutoCADDispatcher)"""
    global _autocad_dispatcher
    _autocad_dispatcher = dispatcher


def set_odoo_util(odoo_util):
    """Set the Odoo utility instance"""
    global _odoo_util
    _odoo_util = odoo_util


def _get_odoo_util():
    """Get or lazy-init Odoo utility"""
    global _odoo_util
    if _odoo_util is not None:
        return _odoo_util
    try:
        from utility.util_odoo import UtilOdoo
        from utility.util_log import UtilLog
        from odoo import sqlite_create_table
        odoo_conn = sqlite_create_table()
        if odoo_conn:
            log_util = UtilLog(None)
            _odoo_util = UtilOdoo(odoo_conn, log_util)
    except Exception as e:
        logger.error(f"Failed to init Odoo util: {e}")
    return _odoo_util


# ==========================================================================
# 5 Odoo-specific tools (registered on the same mcp instance)
# ==========================================================================

@mcp.tool()
def odoo_push_boq(project_id: int = 0) -> Dict[str, Any]:
    """Push BOQ (Bill of Quantities) from AutoCAD to Odoo.

    Extracts TABLE data from all AutoCAD layouts, sends to Odoo import2boq API,
    and writes back the generated header_id/detail_id to the drawing.

    Args:
        project_id: Odoo project ID (0 = use current project from drawing)
    """
    logger.info(f"odoo_push_boq called, project_id={project_id}")
    try:
        if _autocad_dispatcher is None:
            return {"success": False, "error": "AutoCAD not connected"}
        odoo = _get_odoo_util()
        if odoo is None:
            return {"success": False, "error": "Odoo not connected"}

        layouts_data = _autocad_dispatcher.get_layouts_values()
        if not layouts_data:
            return {"success": False, "error": "No TABLE data found in drawing"}

        result = odoo.import2boq(layouts_data)
        if isinstance(result, str):
            return {"success": False, "error": result}
        if not result:
            return {"success": False, "error": "Odoo import2boq returned empty result"}

        _autocad_dispatcher.set_layouts_tables_id(result)

        return {
            "success": True,
            "message": "BOQ pushed and IDs written back",
            "layouts_count": len(layouts_data) if isinstance(layouts_data, list) else 1
        }
    except Exception as e:
        logger.error(f"odoo_push_boq error: {e}")
        return {"success": False, "error": str(e)}


@mcp.tool()
def odoo_create_pr() -> Dict[str, Any]:
    """Create Purchase Requisition from BOQ in Odoo.

    Reads header_ids from AutoCAD TABLEs and calls Odoo boq2pr API.
    """
    logger.info("odoo_create_pr called")
    try:
        if _autocad_dispatcher is None:
            return {"success": False, "error": "AutoCAD not connected"}
        odoo = _get_odoo_util()
        if odoo is None:
            return {"success": False, "error": "Odoo not connected"}

        header_ids = _autocad_dispatcher.get_layouts_header_id_to_pr()
        if not header_ids:
            return {"success": False, "error": "No header_ids found in drawing TABLEs"}

        result = odoo.boq2pr(header_ids)
        return {
            "success": True,
            "message": "Purchase Requisition created",
            "result": result
        }
    except Exception as e:
        logger.error(f"odoo_create_pr error: {e}")
        return {"success": False, "error": str(e)}


@mcp.tool()
def odoo_get_setup(setup_name: str) -> Dict[str, Any]:
    """Get setup/configuration values from Odoo by setup name.

    Args:
        setup_name: Odoo setup record name (e.g. "product_category")
    """
    logger.info(f"odoo_get_setup called, setup_name={setup_name}")
    try:
        odoo = _get_odoo_util()
        if odoo is None:
            return {"success": False, "error": "Odoo not connected"}

        setup = odoo.get_setup(setup_name)
        if isinstance(setup, str):
            return {"success": False, "error": setup}
        return {"success": True, "setup": setup}
    except Exception as e:
        logger.error(f"odoo_get_setup error: {e}")
        return {"success": False, "error": str(e)}


@mcp.tool()
def odoo_get_colors(project_id: int = 0) -> Dict[str, Any]:
    """Get available color options for a project from Odoo.

    Args:
        project_id: Odoo project ID (0 = use the project currently loaded
                    from the AutoCAD drawing)
    """
    logger.info(f"odoo_get_colors called, project_id={project_id}")
    try:
        odoo = _get_odoo_util()
        if odoo is None:
            return {"success": False, "error": "Odoo not connected"}

        if not project_id:
            project_id = getattr(_autocad_dispatcher, 'project_id', None) \
                if _autocad_dispatcher is not None else None
        if not project_id:
            return {"success": False,
                    "error": "No project_id given and none available from the drawing"}

        colors = odoo.get_color(project_id)
        if isinstance(colors, str):
            return {"success": False, "error": colors}
        return {"success": True, "project_id": project_id, "colors": colors}
    except Exception as e:
        logger.error(f"odoo_get_colors error: {e}")
        return {"success": False, "error": str(e)}


@mcp.tool()
def odoo_status() -> Dict[str, Any]:
    """Check connection status for both AutoCAD and Odoo."""
    logger.info("odoo_status called")
    result = {
        "timestamp": datetime.now().isoformat(),
        "autocad": {"connected": False, "mode": None},
        "odoo": {"connected": False}
    }

    if _autocad_dispatcher is not None:
        try:
            result["autocad"]["mode"] = _autocad_dispatcher.mode
            # Avoid calling connected_autocad() which does a sync IPC ping
            # and conflicts with the MCP server's async event loop.
            # Instead, check if the backend was initialized successfully.
            backend = _autocad_dispatcher.active_backend
            if _autocad_dispatcher.mode == "ipc":
                result["autocad"]["connected"] = getattr(backend, '_initialized', False)
            else:
                result["autocad"]["connected"] = getattr(backend, 'acad', None) is not None
        except Exception:
            pass

    odoo = _get_odoo_util()
    if odoo is not None:
        result["odoo"]["connected"] = True

    return result


# ==========================================================================
# Main entry point
# ==========================================================================

def main():
    """Run the MCP server."""
    import argparse

    parser = argparse.ArgumentParser(description="AutoCAD-Odoo MCP Server")
    parser.add_argument("--transport", choices=["stdio", "streamable-http"],
                        default="stdio",
                        help="Transport mode (default: stdio)")
    parser.add_argument("--port", type=int, default=8084,
                        help="Port for HTTP transport (default: 8084)")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler(sys.stderr)],
    )

    logger.info(f"Starting autocad-odoo-mcp server (transport={args.transport})")

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "streamable-http":
        # FastMCP.run() 只接受 (transport, mount_path)；host/port 由 settings 提供
        mcp.settings.host = "127.0.0.1"
        mcp.settings.port = args.port
        mcp.run(transport="streamable-http")


if __name__ == "__main__":
    main()
