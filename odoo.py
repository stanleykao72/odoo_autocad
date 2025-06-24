# -*- coding: utf-8 -*-

import os
import sys
import requests
import simplejson
import yaml
import json
import base64
from requests.exceptions import HTTPError
from pathlib import Path
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

# import YAML for init config
from utility.util_load_yaml_config import LoadYamlConfig

# for sqlite
from models.server import Base as ServerBase
from models.server import Server

# # form
# from autocad.forms.form_main_kivy import FormMain
# form - 使用現代化版本
from forms.form_main_modern import ModernFormMain as FormMain

import logging

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.INFO)

# log to console
c_handler = logging.StreamHandler()

console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
c_handler.setFormatter(console_format)
c_handler.setLevel = logging.DEBUG

_logger.addHandler(c_handler)


def resource_path(relative_path):
    """獲取資源檔案的正確路徑，相容於 PyInstaller 打包"""
    try:
        # PyInstaller 打包後的臨時目錄
        base_path = sys._MEIPASS
    except Exception:
        # 開發環境下的當前目錄
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_app_data_path():
    """獲取應用程式資料目錄路徑"""
    if sys.platform.startswith('win'):
        # Windows: 使用 %APPDATA% 目錄
        app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
        app_dir = os.path.join(app_data, 'OdooAutoCAD')
    else:
        # macOS/Linux: 使用用戶家目錄
        app_dir = os.path.expanduser('~/.odoo_autocad')
    
    # 確保目錄存在
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

def string_to_base64(input_string):
    # 將字串轉換成 UTF-8 編碼的位元組序列
    input_bytes = input_string.encode('utf-8')        
    # 使用 Base64 編碼位元組序列
    base64_bytes = base64.b64encode(input_bytes)
    # 將 Base64 編碼後的位元組序列轉換成字串
    base64_string = base64_bytes.decode('utf-8')
    return base64_string

def sqlite_create_table():
    # 使用應用程式資料目錄來儲存資料庫
    app_data_dir = get_app_data_path()
    db_path = os.path.join(app_data_dir, 'database.db')
    
    _logger.info(f"資料庫路徑: {db_path}")
    
    # 確保資料庫目錄存在
    os.makedirs(os.path.dirname(db_path), exist_ok=True)

    # 建立 Sqlite 引擎
    sqlite_engine = create_engine(f'sqlite:///{db_path}')
    ServerBase.metadata.create_all(sqlite_engine)


    # 使用絕對路徑建立引擎
    Sqlite_session = sessionmaker(bind=sqlite_engine)
    sqlite_session = Sqlite_session()

    # 初始化 server_env
    # server_env = []

    # 尋找 server table 中是否有資料
    sever_env = sqlite_session.query(Server).filter_by(active=True).all()
    # _logger.info(f"first sever_env: {sever_env}")
    # 如果沒有資料，則嘗試從配置檔案載入或使用預設配置
    if not sever_env:
        try:
            # 嘗試從原始配置路徑載入
            config_loader = LoadYamlConfig('c:/odoo/config/token.yaml', 'c:/odoo/config/server.yaml')
            server_cfg, token_cfg = config_loader.load()
        except (FileNotFoundError, Exception) as e:
            _logger.warning(f"無法載入配置檔案: {e}")
            # 使用預設配置
            server_cfg = {
                'host': 'localhost',
                'db_name': 'your_database',
                'url': 'http://localhost:8069/api/v1/swagger.json'
            }
            token_cfg = {
                'token': 'your_token_here'
            }
            _logger.info("使用預設配置，請在首次使用時更新連接設定")
        
        server = Server(host=server_cfg['host'], db_name=server_cfg['db_name'], url=server_cfg['url'], token=token_cfg['token'])
        sqlite_session.add(server)
        sqlite_session.commit()
        new_env = sqlite_session.query(Server).filter_by(active=True).all()
        env = new_env
    else:
        try:
            config_loader = LoadYamlConfig('c:/odoo/config/token.yaml', 'c:/odoo/config/server.yaml')
            server_cfg, token_cfg = config_loader.load()
        except (FileNotFoundError, Exception) as e:
            _logger.warning(f"配置檔案載入失敗，使用資料庫中的現有配置: {e}")
            # 使用現有的資料庫配置
            env = sever_env

        # 如果成功載入配置檔案，則更新資料庫中的資料
        if 'config_loader' in locals() and 'server_cfg' in locals() and 'token_cfg' in locals():
            for env_item in sever_env:
                if env_item.host != server_cfg['host']:
                    env_item.host = server_cfg['host']
                if env_item.db_name != server_cfg['db_name']:
                    env_item.db_name = server_cfg['db_name']
                if env_item.url != server_cfg['url']:
                    env_item.url = server_cfg['url']
                if env_item.token != token_cfg['token']:
                    env_item.token = token_cfg['token']
                sqlite_session.commit()

            new_env = sqlite_session.query(Server).filter_by(active=True).all()
            env = new_env
        else:
            # 使用現有的資料庫配置
            env = sever_env

    if env:
        odoo_env = env[0]
        odoo_conn = {'host': odoo_env.host, 'db_name': odoo_env.db_name, 'url': odoo_env.url, 'token': odoo_env.token}
        return odoo_conn
    else:
        raise ValueError("No server configuration found in the database.")

def connect_to_odoo(odoo_conn):
    host = odoo_conn['host']
    db_name = odoo_conn['db_name']
    url = odoo_conn['url']
    token = odoo_conn['token']
    http_client = RequestsClient()
    http_client.set_basic_auth(host, db_name, token)
    try:
        odoo = SwaggerClient.from_url(url, http_client=http_client)
        # prompt(acaduti, f"與Odoo連線成功\n")
        # print(f"與Odoo連線成功\n")

        basic_string = f'{db_name}:{token}'
        basic_token = string_to_base64(basic_string)
        print(f'basic_token:{basic_token}\n')
        headers = {
            'Authorization': f'Basic {basic_token}'
        }
        print(f'headers:{headers}\n')

        requestOptions = {
            # === bravado config ===
            'headers': headers,
        }

        _logger.info(f"與Odoo連線成功\n")
        return odoo, requestOptions, token
    except requests.exceptions.ConnectionError:
        # print(f"無法與Odoo，通常多試幾次會成功\n")
        _logger.info(f"無法與Odoo，通常多試幾次會成功\n")
        # prompt(acaduti, f"無法與Odoo，通常多試幾次會成功\n")
        return
    except (
        simplejson.errors.JSONDecodeError,      # type: ignore
        yaml.YAMLError,
        HTTPError,
        ):
        # print(
        #     'Invalid swagger file. Please check to make sure the '
        #     'swagger file can be found at: {}.\n'.format(url)
        # )

        _logger.info(f"Invalid swagger file. Please check to make sure the swagger file can be found at: {url}.\n")

        return
    except SwaggerValidationError:
        # print('Invalid swagger format.\n')
        _logger.info(f'Invalid swagger format.')
        return


if __name__ == '__main__':

    odoo_connection = sqlite_create_table()
    # odoo, requestOptions, token = connect_to_odoo(odoo_connection)

    if odoo_connection:
        # print(f"odoo: {odoo}")
        # print(f"requestOptions: {requestOptions}")
        # print(f"token: {token}")
        # prompt(acaduti, f"odoo: {odoo}")
        # prompt(acaduti, f"requestOptions: {requestOptions}")
        # prompt(acaduti, f"token: {token}")

        # 建立 FormMain
        # form_main = FormMain(odoo=odoo, requestOptions=requestOptions, token=token, odoo_connection=odoo_connection)
        form_main = FormMain(odoo_connection=odoo_connection)
        form_main.mainloop()
