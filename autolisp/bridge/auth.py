"""
auth.py — Authentication management for Odoo API.

Supports token-based auth: BasicAuth(db_name, token)
Compatible with the existing server_prod.yaml + token.yaml setup.
"""

import base64
from requests.auth import HTTPBasicAuth

from config import BridgeConfig


class AuthManager:
    """Manages authentication for Odoo API requests."""

    def __init__(self, config: BridgeConfig):
        self.config = config

    def get_auth(self):
        """Returns HTTPBasicAuth(db_name, token)."""
        return HTTPBasicAuth(
            self.config.username,
            self.config.password
        )

    def get_basic_header(self) -> dict:
        """Returns Authorization header with Base64-encoded db:token.
        Matches the format used by the main Python app (util_odoo.py)."""
        basic_string = f'{self.config.db}:{self.config.token}'
        basic_bytes = base64.b64encode(basic_string.encode('utf-8'))
        basic_token = basic_bytes.decode('utf-8')
        return {'Authorization': f'Basic {basic_token}'}
