# -*- coding: utf-8 -*-

import os
import sys
import requests
import simplejson
import yaml
import json
import base64
import argparse
import signal
import time
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


def start_mcp_server_only(args):
    """啟動純MCP伺服器模式 (無GUI)"""
    try:
        from ai_assistant.mcp_server_manager import MCPServerManager
        from utility.util_log import UtilLog
        from utility.util_autocad import UtilAutoCAD
        from utility.util_odoo import UtilOdoo
        
        _logger.info("啟動MCP伺服器模式...")
        
        # 建立資料庫連接
        odoo_connection = sqlite_create_table()
        if not odoo_connection:
            _logger.error("無法建立Odoo連接配置")
            return 1
        
        # 建立工具實例 (不需要GUI組件)
        log_util = UtilLog(None)  # None因為沒有GUI
        odoo_util = UtilOdoo(odoo_connection, log_util)
        autocad_util = UtilAutoCAD(odoo_util, log_util)
        
        # 建立MCP伺服器管理器
        server_manager = MCPServerManager(
            autocad_util=autocad_util,
            odoo_util=odoo_util,
            log_util=log_util,
            tcp_port=args.mcp_port,
            pipe_name=args.mcp_pipe
        )
        
        # 設置信號處理程序以正常關閉
        def signal_handler(signum, frame):
            _logger.info("收到停止信號，正在關閉MCP伺服器...")
            server_manager.stop_all_servers()
            sys.exit(0)
        
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
        
        # 啟動伺服器
        server_manager.start_all_servers()
        _logger.info(f"MCP伺服器已啟動 - TCP:{args.mcp_port}, Pipe:{args.mcp_pipe}")
        
        # 保持運行
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            _logger.info("收到鍵盤中斷，正在關閉...")
        finally:
            server_manager.stop_all_servers()
        
        return 0
        
    except Exception as e:
        _logger.error(f"MCP伺服器啟動失敗: {e}")
        return 1


def start_mcp_autocad_server(args):
    """啟動 MCP AutoCAD 伺服器 (mcp_server_autocad.py, stdio)"""
    try:
        import mcp_server_autocad
        _logger.info("Starting MCP AutoCAD server (stdio)...")
        mcp_server_autocad.main()
        return 0
    except Exception as e:
        _logger.error(f"MCP AutoCAD server failed: {e}")
        return 1


def start_gui_application(enable_mcp=False, mcp_args=None, autocad_mode="com"):
    """啟動GUI應用程式 (可選擇性啟用MCP)"""
    odoo_connection = sqlite_create_table()

    if odoo_connection:
        form_main = FormMain(odoo_connection=odoo_connection, autocad_mode=autocad_mode)
        
        # 如果啟用MCP，自動啟動MCP伺服器
        if enable_mcp and hasattr(form_main, 'initialize_mcp_server_manager'):
            try:
                form_main.initialize_mcp_server_manager()
                if form_main.mcp_server_manager:
                    form_main.mcp_server_manager.start_server()
                    _logger.info("GUI模式下自動啟動MCP SSE伺服器")
            except Exception as e:
                _logger.warning(f"GUI模式下自動啟動MCP伺服器失敗: {e}")
        
        form_main.mainloop()


def main():
    """主程式入口點"""
    parser = argparse.ArgumentParser(
        description='Odoo AutoCAD Integration with MCP Support',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例用法:
  python odoo.py                           # 啟動GUI模式
  python odoo.py --enable-mcp              # 啟動GUI並自動啟動MCP伺服器
  python odoo.py --mcp-server               # 純MCP伺服器模式 (無GUI)
  python odoo.py --mcp-server --mcp-port 8001  # 自訂端口的MCP伺服器
        """
    )
    
    # MCP 伺服器相關參數
    parser.add_argument('--mcp-server', action='store_true',
                       help='啟動純MCP伺服器模式 (無GUI)')
    parser.add_argument('--mcp-port', type=int, default=8000,
                       help='MCP TCP伺服器端口 (預設: 8000)')
    parser.add_argument('--mcp-pipe', type=str, 
                       default=r'\\.\pipe\odoo_autocad_mcp',
                       help='MCP Named Pipe名稱')
    parser.add_argument('--enable-mcp', action='store_true',
                       help='在GUI模式下啟用MCP伺服器')
    parser.add_argument('--autocad-mode', choices=['com', 'ipc'], default='com',
                       help='AutoCAD connection mode: com (Full AutoCAD) or ipc (LT 2024+)')
    parser.add_argument('--mcp-autocad', action='store_true',
                       help='啟動MCP模式操作AutoCAD（使用 mcp_server_autocad.py）')

    args = parser.parse_args()
    
    if args.mcp_autocad:
        # MCP AutoCAD 模式 (stdio, uses mcp_server_autocad.py)
        return start_mcp_autocad_server(args)
    elif args.mcp_server:
        # 純MCP伺服器模式
        return start_mcp_server_only(args)
    else:
        # GUI模式
        start_gui_application(
            enable_mcp=args.enable_mcp,
            mcp_args=args,
            autocad_mode=args.autocad_mode
        )
        return 0


if __name__ == '__main__':
    sys.exit(main())
