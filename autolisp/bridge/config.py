"""
config.py — Configuration reader for the bridge.

Supports two modes:
  1. YAML mode (production): reads server_prod.yaml + token.yaml
  2. INI mode (legacy): reads bridge.ini

YAML mode uses the same config files as the main Python Tkinter app.
Auth: BasicAuth(db_name, token)
"""

import os
import configparser

try:
    import yaml
except ImportError:
    yaml = None


class BridgeConfig:
    """Configuration loaded from YAML or INI files."""

    def __init__(self, config_path: str = None,
                 server_yaml: str = None, token_yaml: str = None):
        """
        Args:
            config_path: Path to bridge.ini (INI mode)
            server_yaml: Path to server_prod.yaml (YAML mode)
            token_yaml:  Path to token.yaml (YAML mode)
        """
        # Defaults
        self.url = ''
        self.host = ''
        self.db = ''
        self.swagger_url = ''
        self.auth_method = 'token'
        self.username = ''
        self.password = ''
        self.token = ''
        self.timeout = 30
        self.debug = False
        self.verify_ssl = True
        self.temp_dir = os.path.join(os.environ.get('TEMP', '/tmp'), 'odoo_bridge')
        self.log_file = os.path.join(self.temp_dir, 'bridge.log')

        # Try YAML mode first
        if server_yaml or token_yaml:
            self._load_yaml(server_yaml, token_yaml)
        elif config_path and config_path.endswith('.yaml'):
            # Single YAML path = server config, try to find token from it
            self._load_yaml(config_path, None)
        else:
            # Fallback: try auto-detect YAML in standard locations
            yaml_loaded = self._try_auto_detect_yaml(config_path)
            if not yaml_loaded:
                self._load_ini(config_path)

    def _try_auto_detect_yaml(self, config_path: str = None) -> bool:
        """Try to find server_prod.yaml + token.yaml in standard locations."""
        if yaml is None:
            return False

        search_dirs = []
        if config_path and os.path.isfile(config_path):
            search_dirs.append(os.path.dirname(os.path.abspath(config_path)))
        # Standard location: ../../config/ relative to bridge/
        bridge_dir = os.path.dirname(os.path.abspath(__file__))
        search_dirs.append(os.path.join(bridge_dir, '..', '..', 'config'))
        search_dirs.append(os.path.join(bridge_dir, '..', 'config'))

        for d in search_dirs:
            server_path = os.path.join(d, 'server_prod.yaml')
            if os.path.exists(server_path):
                self._load_yaml(server_path, None)
                return True
        return False

    def _load_yaml(self, server_yaml: str = None, token_yaml: str = None):
        """Load configuration from YAML files."""
        if yaml is None:
            raise ImportError("PyYAML is required for YAML config. Install: pip install pyyaml")

        server_data = {}
        token_data = {}

        # Load server config
        if server_yaml and os.path.exists(server_yaml):
            with open(server_yaml, 'r', encoding='utf-8') as f:
                raw = yaml.safe_load(f)
                server_data = raw.get('server', raw) if raw else {}

        # Resolve token.yaml path
        if not token_yaml and server_data.get('token_file'):
            token_yaml = server_data['token_file']
            # Handle relative paths
            if not os.path.isabs(token_yaml) and server_yaml:
                token_yaml = os.path.join(os.path.dirname(server_yaml), token_yaml)

        # Load token config
        if token_yaml and os.path.exists(token_yaml):
            with open(token_yaml, 'r', encoding='utf-8') as f:
                raw = yaml.safe_load(f)
                token_data = raw.get('user', raw) if raw else {}

        # Map to config properties
        self.host = server_data.get('host', '')
        self.db = server_data.get('db_name', '')
        self.swagger_url = server_data.get('url', '')
        self.url = f"https://{self.host}" if self.host else ''
        self.token = token_data.get('token', '')
        self.auth_method = 'token'

        # For BasicAuth: username = db_name, password = token
        self.username = self.db
        self.password = self.token

    def _load_ini(self, config_path: str = None):
        """Load configuration from INI file (legacy mode)."""
        parser = configparser.ConfigParser()

        if config_path and os.path.exists(config_path):
            ini_path = config_path
        else:
            candidates = [
                os.path.join(os.path.dirname(__file__), '..', 'config', 'bridge.ini'),
                os.path.join(os.path.dirname(__file__), 'bridge.ini'),
            ]
            ini_path = None
            for c in candidates:
                if os.path.exists(c):
                    ini_path = os.path.abspath(c)
                    break

        if ini_path:
            parser.read(ini_path, encoding='utf-8')

        def _get(section, key, default=''):
            try:
                return parser.get(section, key)
            except (configparser.NoSectionError, configparser.NoOptionError):
                return default

        self.url = _get('odoo', 'url', 'https://odoo.example.com')
        self.db = _get('odoo', 'db', '')
        self.auth_method = _get('odoo', 'auth_method', 'basic')
        self.username = _get('credentials', 'username', '')
        self.password = _get('credentials', 'password', '')
        self.temp_dir = os.path.expandvars(
            _get('paths', 'temp_dir', self.temp_dir))
        self.log_file = os.path.expandvars(
            _get('paths', 'log_file', self.log_file))
        self.timeout = int(_get('options', 'timeout', '30'))
        self.debug = _get('options', 'debug', 'false').lower() == 'true'
        self.verify_ssl = _get('options', 'verify_ssl', 'true').lower() == 'true'
