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
# form
from forms.form_main import FormMain

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
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def string_to_base64(input_string):
    # 將字串轉換成 UTF-8 編碼的位元組序列
    input_bytes = input_string.encode('utf-8')        
    # 使用 Base64 編碼位元組序列
    base64_bytes = base64.b64encode(input_bytes)
    # 將 Base64 編碼後的位元組序列轉換成字串
    base64_string = base64_bytes.decode('utf-8')
    return base64_string

def sqlite_create_table():
    # 獲取當前檔案的目錄
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 設定資料庫路徑
    db_path = os.path.join(resource_path(current_dir), 'db', 'database.db')
    # print('db_path: %s' % db_path)

    # 建立 Sqlite 引擎
    sqlite_engine = create_engine(f'sqlite:///{db_path}')
    ServerBase.metadata.create_all(sqlite_engine)


    # 使用絕對路徑建立引擎
    Sqlite_session = sessionmaker(bind=sqlite_engine)
    sqlite_session = Sqlite_session()

    # 初始化 server_env
    server_env = []

    # 尋找 server table 中是否有資料
    sever_env = sqlite_session.query(Server).filter_by(active=True).all()
    # _logger.info(f"first sever_env: {sever_env}")
    # 如果沒有資料，則從LoadYamlConfig取得server_cfg, token_cfg 新增一筆資料
    if not sever_env:
        config_loader = LoadYamlConfig('c:/odoo/config/token.yaml', 'c:/odoo/config/server.yaml')
        server_cfg, token_cfg = config_loader.load()
        server = Server(host=server_cfg['host'], db_name=server_cfg['db_name'], url=server_cfg['url'], token=token_cfg['token'])
        sqlite_session.add(server)
        sqlite_session.commit()
        new_env = sqlite_session.query(Server).filter_by(active=True).all()
        # _logger.info(f"second server_env: {server_env}")
        env = new_env
    else:
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
    odoo, requestOptions, token = connect_to_odoo(odoo_connection)

    if odoo:
        # print(f"odoo: {odoo}")
        # print(f"requestOptions: {requestOptions}")
        # print(f"token: {token}")
        # prompt(acaduti, f"odoo: {odoo}")
        # prompt(acaduti, f"requestOptions: {requestOptions}")
        # prompt(acaduti, f"token: {token}")

        # 建立 FormMain
        form_main = FormMain(odoo=odoo, requestOptions=requestOptions, token=token, odoo_connection=odoo_connection)
        form_main.mainloop()
