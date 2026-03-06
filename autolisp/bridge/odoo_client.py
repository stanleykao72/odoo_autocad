"""
odoo_client.py — Odoo Swagger/REST API client.

Uses requests + BasicAuth to communicate with Odoo's custom endpoints.
API base path: /api/v1/boq_import_api/
All business methods use PATCH on: job.working.plan.boq/call/{method_name}
Auth: BasicAuth(db_name, token)
"""

import requests
import logging
from typing import Any, Optional

from config import BridgeConfig
from auth import AuthManager

logger = logging.getLogger(__name__)

# Swagger API base path (matches Odoo module registration)
API_BASE = "/api/v1/boq_import_api"
# Model path for all BOQ-related method calls
BOQ_MODEL = "job.working.plan.boq"


class OdooClient:
    """Client for Odoo REST/Swagger API communication."""

    def __init__(self, config: BridgeConfig):
        self.config = config
        self.auth = AuthManager(config)
        self.session = requests.Session()
        self.session.verify = config.verify_ssl
        self.base_url = f"https://{config.host}" if config.host else config.url.rstrip('/')
        self.timeout = config.timeout
        self.user_token = config.token

    def _call_method_url(self, method_name: str) -> str:
        """Build URL for PATCH /api/v1/boq_import_api/job.working.plan.boq/call/{method_name}"""
        return f"{self.base_url}{API_BASE}/{BOQ_MODEL}/call/{method_name}"

    def _request(self, method: str, url: str, **kwargs) -> dict:
        """Make an authenticated request."""
        kwargs.setdefault('timeout', self.timeout)
        # Use BasicAuth(db_name, token)
        kwargs['auth'] = self.auth.get_auth()
        kwargs.setdefault('headers', {})
        kwargs['headers']['Content-Type'] = 'application/json'
        # Also add Authorization header (base64 db:token) for compatibility
        kwargs['headers'].update(self.auth.get_basic_header())

        try:
            logger.info(f"Request: {method} {url}")
            resp = self.session.request(method, url, **kwargs)
            resp.raise_for_status()
            return {"success": True, "data": resp.json()}
        except requests.ConnectionError as e:
            logger.error(f"Connection error: {e}")
            return {"success": False, "error_code": "CONNECTION_ERROR",
                    "message": f"Cannot connect to {self.base_url}"}
        except requests.Timeout:
            return {"success": False, "error_code": "TIMEOUT",
                    "message": f"Request timed out after {self.timeout}s"}
        except requests.HTTPError as e:
            status = e.response.status_code if e.response else 0
            if status == 401:
                return {"success": False, "error_code": "AUTH_FAILED",
                        "message": "Authentication failed: invalid credentials"}
            body = ""
            try:
                body = e.response.json().get("message", str(e))
            except Exception:
                body = str(e)
            return {"success": False, "error_code": f"HTTP_{status}",
                    "message": body}
        except Exception as e:
            logger.error(f"Unexpected error: {e}")
            return {"success": False, "error_code": "UNKNOWN_ERROR",
                    "message": str(e)}

    def _call_method(self, method_name: str, args: list = None,
                     kwargs_extra: dict = None) -> dict:
        """Call a method on job.working.plan.boq model via PATCH.
        Matches the pattern used by the main Python app (util_odoo.py):
          PATCH .../job.working.plan.boq/call/{method_name}
          Body: {"args": [...], "kwargs": {"user_token": "..."}, "context": {}}
        """
        url = self._call_method_url(method_name)
        body = {
            "args": args or [],
            "kwargs": {"user_token": self.user_token},
            "context": {}
        }
        if kwargs_extra:
            body["kwargs"].update(kwargs_extra)
        return self._request("PATCH", url, json=body)

    # ============================================================
    # API Methods
    # ============================================================

    def test_connection(self) -> dict:
        """Test connection by fetching the swagger spec or calling a lightweight method."""
        # No dedicated test_connection endpoint in swagger.
        # Use GET on the model list as a connectivity check.
        url = f"{self.base_url}{API_BASE}/{BOQ_MODEL}"
        result = self._request("GET", url, params={"limit": 1})
        if result.get("success"):
            return {"success": True, "data": {
                "server": self.base_url,
                "database": self.config.db,
                "status": "connected"
            }}
        return result

    def get_project(self, pr_no: str) -> dict:
        """Get project info by PR number."""
        return self._call_method("get_project_v2",
                                 args=[[['name', '=', pr_no]]])

    def get_products(self) -> dict:
        """Get product list."""
        return self._call_method("get_product_v2",
                                 args=[[('categ_id', 'child_of', 27),
                                        ('active', '=', True)]])

    def get_setup(self, setup_name: str = None) -> dict:
        """Get setup values (spec, catalog, operation_flow, surface_treatment)."""
        args = [[('setup_name', '=', setup_name)]] if setup_name else []
        return self._call_method("get_setup_v2", args=args)

    def get_colors(self, project_id: Optional[int] = None) -> dict:
        """Get color list, optionally filtered by project."""
        args = [[('job_project_id', '=', project_id)]] if project_id else []
        return self._call_method("get_color_v2", args=args)

    def import_to_boq(self, data: dict) -> dict:
        """Import BOQ data to Odoo."""
        return self._call_method("import2boq_v2", args=[data])

    def boq_to_pr(self, header_ids: Any) -> dict:
        """Convert BOQ headers to Purchase Requisition."""
        if isinstance(header_ids, str):
            import json
            header_ids = json.loads(header_ids)
        return self._call_method("boq2pr_v2", args=[header_ids])
