# -*- coding: utf-8 -*-
"""
MCP Server for AutoCAD-Odoo Integration — 13 tools (8 common + 5 Odoo)

Extends autocad-mcp's 8 tools with 5 Odoo-specific tools.
Supports stdio and streamable-http transports.

Tools (8 from autocad-mcp):
  drawing, entity, layer, block, annotation, pid, view, system

Tools (5 Odoo-specific):
  odoo_push_boq, odoo_create_pr, odoo_get_setup, odoo_get_colors, odoo_status
"""

import os
import sys
import logging
import json
from datetime import datetime
from typing import Dict, Any

from mcp.server.fastmcp import FastMCP

# Add autocad-mcp to path (for IPC backend usage, not tool import)
_autocad_mcp_src = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "libs", "autocad-mcp", "src"
)
if _autocad_mcp_src not in sys.path:
    sys.path.insert(0, _autocad_mcp_src)

logger = logging.getLogger(__name__)

# Create our MCP server
mcp = FastMCP("autocad-odoo-mcp")

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
# Note: autocad-mcp's 8 upstream tools (drawing, entity, layer, block,
# annotation, pid, view, system) are NOT imported here due to FastMCP
# type annotation incompatibility. They remain available via:
#   1. Running autocad-mcp's own server separately
#   2. Using execute_lisp through the IPC backend
#   3. Using our Odoo tools which call the IPC backend internally
# ==========================================================================


# ==========================================================================
# 5 Odoo-specific tools
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

        # Step 1: Extract TABLE data from AutoCAD
        layouts_data = _autocad_dispatcher.get_layouts_values()
        if not layouts_data:
            return {"success": False, "error": "No TABLE data found in drawing"}

        # Step 2: Push to Odoo
        result = odoo.import2boq(layouts_data)
        if not result:
            return {"success": False, "error": "Odoo import2boq returned empty result"}

        # Step 3: Write IDs back to AutoCAD
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
def odoo_get_setup(project_id: int = 0) -> Dict[str, Any]:
    """Get setup/configuration values from Odoo (products, specs, etc.)

    Args:
        project_id: Odoo project ID
    """
    logger.info(f"odoo_get_setup called, project_id={project_id}")
    try:
        odoo = _get_odoo_util()
        if odoo is None:
            return {"success": False, "error": "Odoo not connected"}

        setup = odoo.get_setup(project_id)
        return {"success": True, "setup": setup}
    except Exception as e:
        logger.error(f"odoo_get_setup error: {e}")
        return {"success": False, "error": str(e)}


@mcp.tool()
def odoo_get_colors() -> Dict[str, Any]:
    """Get available color options from Odoo."""
    logger.info("odoo_get_colors called")
    try:
        odoo = _get_odoo_util()
        if odoo is None:
            return {"success": False, "error": "Odoo not connected"}

        colors = odoo.get_color()
        return {"success": True, "colors": colors}
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
            result["autocad"]["connected"] = _autocad_dispatcher.connected_autocad()
            result["autocad"]["mode"] = _autocad_dispatcher.mode
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
    """Run the MCP server on stdio transport."""
    import argparse

    parser = argparse.ArgumentParser(description="AutoCAD-Odoo MCP Server")
    parser.add_argument("--transport", choices=["stdio", "sse"], default="stdio",
                        help="Transport mode (default: stdio)")
    parser.add_argument("--port", type=int, default=8084,
                        help="Port for SSE transport (default: 8084)")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler(sys.stderr)],
    )

    logger.info(f"Starting autocad-odoo-mcp server (transport={args.transport})")

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "sse":
        mcp.run(transport="sse", sse_params={"port": args.port})


if __name__ == "__main__":
    main()
