# -*- coding: utf-8 -*-

import yaml
import logging

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

        # print(f'token_config:{token_config}\n')
        _logger.info(f'token_config:{token_config}')
        token_cfg = token_config['user']
        # print(f'token_cfg:{token_cfg}\n')
        _logger.info(f'token_cfg:{token_cfg}')

        if 'token' in token_cfg:
            pass
        else:
            # prompt(acaduti, 'token.yaml設定中沒有 token\n')
            return "token.yaml設定中沒有 token\n"
        
        if 'server_file' in token_cfg:
            with open(f"{token_cfg['server_file']}", 'r') as f:
                server_config = yaml.load(f, Loader=yaml.FullLoader)
        
            # print(server_config)
            # _logger.info(f'server_config: {server_config}')
            server_cfg = server_config['server']
            # print(f'server_cfg: {server_cfg}\n')
            _logger.info(f'server_cfg: {server_cfg}')
            # return server_cfg
        else:
            # prompt(acaduti, 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n')
            return 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n'

        return server_cfg, token_cfg