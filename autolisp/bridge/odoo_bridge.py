#!/usr/bin/env python3
"""
odoo_bridge.py — CLI entry point for AutoLISP-Odoo communication.

Usage:
    odoo_bridge.py <action> <request.json> <response.json> [--config bridge.ini]

Actions:
    test_connection   Test Odoo connection
    get_project       Get project info (requires pr_no param)
    get_products      Get product list
    get_setup         Get setup values
    get_colors        Get color list (requires project_id param)
    import_to_boq     Import BOQ data
    boq_to_pr         Convert BOQ to PR (requires header_ids param)
"""

import sys
import json
import argparse
import traceback
import logging

from config import BridgeConfig
from odoo_client import OdooClient

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)


def read_request(filepath: str) -> dict:
    """Read and parse request JSON file."""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def write_response(filepath: str, data: dict):
    """Write response JSON file."""
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def success_response(data=None, message="OK"):
    """Build a success response."""
    resp = {"success": True, "message": message}
    if data is not None:
        resp["data"] = data
    return resp


def error_response(error_code: str, message: str):
    """Build an error response."""
    return {
        "success": False,
        "error_code": error_code,
        "message": message
    }


def dispatch(action: str, params: dict, client: OdooClient) -> dict:
    """Dispatch action to appropriate handler."""
    handlers = {
        "test_connection": handle_test_connection,
        "get_project": handle_get_project,
        "get_products": handle_get_products,
        "get_setup": handle_get_setup,
        "get_colors": handle_get_colors,
        "import_to_boq": handle_import_to_boq,
        "boq_to_pr": handle_boq_to_pr,
    }

    handler = handlers.get(action)
    if not handler:
        return error_response("UNKNOWN_ACTION", f"Unknown action: {action}")

    return handler(params, client)


def handle_test_connection(params: dict, client: OdooClient) -> dict:
    """Test Odoo connection."""
    result = client.test_connection()
    if result.get("success"):
        return success_response(result.get("data"), "Connection successful")
    return error_response("CONNECTION_FAILED", result.get("message", "Connection failed"))


def handle_get_project(params: dict, client: OdooClient) -> dict:
    """Get project info by PR number."""
    pr_no = params.get("pr_no")
    if not pr_no:
        return error_response("MISSING_PARAM", "pr_no is required")
    result = client.get_project(pr_no)
    if result.get("success"):
        return success_response(result["data"])
    return error_response(result.get("error_code", "API_ERROR"), result.get("message", "Failed"))


def handle_get_products(params: dict, client: OdooClient) -> dict:
    """Get product list."""
    result = client.get_products()
    if result.get("success"):
        return success_response(result["data"])
    return error_response(result.get("error_code", "API_ERROR"), result.get("message", "Failed"))


def handle_get_setup(params: dict, client: OdooClient) -> dict:
    """Get setup values."""
    result = client.get_setup()
    if result.get("success"):
        return success_response(result["data"])
    return error_response(result.get("error_code", "API_ERROR"), result.get("message", "Failed"))


def handle_get_colors(params: dict, client: OdooClient) -> dict:
    """Get colors for a project."""
    project_id = params.get("project_id")
    result = client.get_colors(project_id)
    if result.get("success"):
        return success_response(result["data"])
    return error_response(result.get("error_code", "API_ERROR"), result.get("message", "Failed"))


def handle_import_to_boq(params: dict, client: OdooClient) -> dict:
    """Import BOQ data to Odoo."""
    result = client.import_to_boq(params)
    if result.get("success"):
        return success_response(result["data"], "Successfully imported to BOQ")
    return error_response(result.get("error_code", "API_ERROR"), result.get("message", "Failed"))


def handle_boq_to_pr(params: dict, client: OdooClient) -> dict:
    """Convert BOQ to Purchase Requisition."""
    header_ids = params.get("header_ids")
    if not header_ids:
        return error_response("MISSING_PARAM", "header_ids is required")
    result = client.boq_to_pr(header_ids)
    if result.get("success"):
        return success_response(result["data"], "Successfully transferred to PR")
    return error_response(result.get("error_code", "API_ERROR"), result.get("message", "Failed"))


def main():
    parser = argparse.ArgumentParser(description="Odoo Bridge CLI")
    parser.add_argument("action", help="Action to perform")
    parser.add_argument("request_file", help="Path to request JSON file")
    parser.add_argument("response_file", help="Path to response JSON file")
    parser.add_argument("--config", default=None, help="Path to bridge.ini config file")
    parser.add_argument("--server-config", default=None,
                        help="Path to server_prod.yaml")
    parser.add_argument("--token-config", default=None,
                        help="Path to token.yaml")

    args = parser.parse_args()

    try:
        # Load config (YAML takes priority over INI)
        cfg = BridgeConfig(
            config_path=args.config,
            server_yaml=args.server_config,
            token_yaml=args.token_config
        )

        # Read request
        request = read_request(args.request_file)
        params = request.get("params", {})
        # If params is a list with one dict, unwrap it
        if isinstance(params, list) and len(params) == 1 and isinstance(params[0], dict):
            params = params[0]
        elif isinstance(params, list) and len(params) == 0:
            params = {}

        # Create client and dispatch
        client = OdooClient(cfg)
        response = dispatch(args.action, params, client)

    except FileNotFoundError as e:
        response = error_response("FILE_NOT_FOUND", str(e))
    except json.JSONDecodeError as e:
        response = error_response("INVALID_JSON", f"Invalid JSON in request: {e}")
    except Exception as e:
        logger.error(traceback.format_exc())
        response = error_response("INTERNAL_ERROR", str(e))

    # Write response
    try:
        write_response(args.response_file, response)
    except Exception as e:
        # Last resort: try to write a minimal error
        try:
            with open(args.response_file, 'w') as f:
                json.dump(error_response("WRITE_ERROR", str(e)), f)
        except Exception:
            sys.exit(1)


if __name__ == "__main__":
    main()
