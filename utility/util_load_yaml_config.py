# -*- coding: utf-8 -*-

import yaml
import logging

from utility.util_secrets import mask_url_token

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.INFO)

# log to console
c_handler = logging.StreamHandler()

console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
c_handler.setFormatter(console_format)
c_handler.setLevel = logging.DEBUG

_logger.addHandler(c_handler)


class LoadYamlConfig:
    def __init__(self, token_yaml_file, server_yaml_file):
        self.token_yaml_file = token_yaml_file
        self.server_yaml_file = server_yaml_file
        self.load()

    def load(self):
        with open(self.token_yaml_file, 'r') as stream:
            token_config = yaml.load(stream, Loader=yaml.FullLoader)

        token_cfg = token_config['user']
        # 不記錄 token 明文：只列出有哪些設定鍵
        _logger.info(f'token_cfg keys: {sorted(token_cfg.keys())}')

        if 'token' in token_cfg:
            pass
        else:
            # prompt(acaduti, 'token.yaml設定中沒有 token\n')
            return "token.yaml設定中沒有 token\n"
        
        if 'server_file' in token_cfg:
            with open(f"{token_cfg['server_file']}", 'r') as f:
                server_config = yaml.load(f, Loader=yaml.FullLoader)
        
            server_cfg = server_config['server']
            # 不可直接記錄整個 server_cfg：url 的 query string 內含 API token。
            _logger.info(
                f"server_cfg: host={server_cfg.get('host')}, "
                f"db_name={server_cfg.get('db_name')}, "
                f"url={mask_url_token(server_cfg.get('url'))}")
        else:
            # prompt(acaduti, 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n')
            return 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n'

        return server_cfg, token_cfg