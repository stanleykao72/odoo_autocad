# -*- coding: utf-8 -*-

import sys
import pythoncom
import yaml
import json
import base64
import requests
import simplejson

from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from bravado.exception import HTTPError
from swagger_spec_validator.common import SwaggerValidationError    # type: ignore

import logging

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.INFO)

# log to console
c_handler = logging.StreamHandler()

console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
c_handler.setFormatter(console_format)
c_handler.setLevel = logging.DEBUG

_logger.addHandler(c_handler)


class ComServer:
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER
    _reg_clsid_ = "{F4149E66-4186-4CD9-A677-D24495D50FC6}"
    _reg_desc_ = "Python COM Server"
    _reg_progid_ = "Python.ComServer"
    _public_methods_ = ['Hello', 'yaml_to_json', 'string_to_base64', 'odoo_connection', 'import2boq', 'boq2pr', 'get_project', 'get_product', 'get_setup', 'get_color', 'get_server_config']
    _public_attrs_ = ['softspace', 'noCalls', 'odoo', 'user_token']
    _readonly_attrs_ = ['noCalls']
    # for Python 3.7+
    _reg_verprogid_ = "Python.ComServer.1"
    _reg_class_spec_ = __name__ + ".ComServer"

    def __init__(self):
        self.softspace = 1
        self.noCalls = 0
        self.odoo = False
        self.user_token = False

    def Hello(self, who):
        self.noCalls = self.noCalls + 1
        # insert "softspace" number of spaces
        return "Hello" + " " * self.softspace + str(who)

    def yaml_to_json(self, yaml_file):
        with open(yaml_file, 'r') as file:
            data = yaml.safe_load(file)
        return json.dumps(data)

    def string_to_base64(self, input_string):
        # 將字串轉換成 UTF-8 編碼的位元組序列
        input_bytes = input_string.encode('utf-8')        
        # 使用 Base64 編碼位元組序列
        base64_bytes = base64.b64encode(input_bytes)
        # 將 Base64 編碼後的位元組序列轉換成字串
        base64_string = base64_bytes.decode('utf-8')
        return base64_string

    def get_server_config(self):
        # read config file
        with open('c:/odoo/config/token.yaml', 'r') as f:
            token_config = yaml.load(f, Loader=yaml.FullLoader)

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
            return server_cfg
        else:
            # prompt(acaduti, 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n')
            return 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n'

    def odoo_connection(self):
        # read config file
        with open('c:/odoo/config/token.yaml', 'r') as f:
            token_config = yaml.load(f, Loader=yaml.FullLoader)

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
        else:
            # prompt(acaduti, 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n')
            return 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n'

        host = server_cfg['host']  #'odoo-esmith-1124-stage-6571675.dev.odoo.com'
        db_name = server_cfg['db_name'] #'odoo-esmith-1124-stage-6571675'
        user_token = token_cfg['token'] #'6d4bead3-8c1a-46b4-a399-7e57535b85d9'
        self.user_token = user_token
        url = server_cfg['url'] #'https://odoo-esmith-1124-stage-6571675.dev.odoo.com/api/v1/boq_import_api/swagger.json?token=34dba8ba-cf29-4ac7-a2b7-e64b7ac7bae6&db=odoo-esmith-1124-stage-6571675'
        http_client = RequestsClient()
        http_client.set_basic_auth(host, db_name, user_token)

        try:
            odoo = SwaggerClient.from_url(
                url,
                http_client=http_client
            )
            # prompt(acaduti, f"與Odoo連線成功\n")
            # print(f"與Odoo連線成功\n")
            _logger.info(f"與Odoo連線成功\n")
            self.odoo = odoo
            import_return_str = json.dumps(server_cfg, ensure_ascii=False).encode('utf8').decode()
            return import_return_str
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

    def import2boq(self, header_json):
        # print(f'user_token:{self.user_token}\n')
        # print(f'header_json:{header_json}')
        # print(f'odoo:{self.odoo}\n')

        header_dict = json.loads(header_json)
        # print(f'header_dict:{header_dict}')
        import_return_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="import2boq_v2",
            body={
            "args": [header_dict],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in import_return_list:
            error_message = import_return_list.get('error_message')
            _logger.info(f'error_message: {error_message}')
            import_return_str = json.dumps(import_return_list, ensure_ascii=False).encode('utf8').decode()
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str
        else:
            import_return_str = json.dumps(import_return_list, ensure_ascii=False).encode('utf8').decode()
            # print(f'import_return_str:{import_return_str}')
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str

    def boq2pr(self, header_json):
        # print(f'user_token:{self.user_token}\n')
        # print(f'header_json:{header_json}')
        # print(f'odoo:{self.odoo}\n')

        header_dict = json.loads(header_json)
        # print(f'header_dict:{header_dict}')
        import_return_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="boq2pr_v2",
            body={
            "args": [header_dict],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in import_return_list:
            error_message = import_return_list.get('error_message')
            _logger.info(f'error_message: {error_message}')
            return error_message
        else:
            import_return_str = json.dumps(import_return_list, ensure_ascii=False).encode('utf8').decode()
            # print(f'import_return_str:{import_return_str}')
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str

    def get_project(self, pr_no):

        project_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_project_v2",
            body={
            "args": [[('name', '=', pr_no)]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in project_list:
            error_message = project_list.get('error_message')
            _logger.info(f'error_message: {error_message}')
            return error_message
        else:
            import_return_str = json.dumps(project_list, ensure_ascii=False).encode('utf8').decode()
            # print(f'import_return_str:{import_return_str}')
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str

    def get_product(self):

        product_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_product_v2",
            body={
            "args": [[('categ_id', 'child_of', 27), ('active', '=', True)]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in product_list:
            error_message = product_list.get('error_message')
            _logger.info(f'error_message: {error_message}')
            return error_message
        else:
            import_return_str = json.dumps(product_list, ensure_ascii=False).encode('utf8').decode()
            # print(f'import_return_str:{import_return_str}')
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str

    def get_setup(self):

        setup_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_setup_v2",
            body={
            "args": [[]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in setup_list:
            error_message = setup_list.get('error_message')
            _logger.info(f'error_message: {error_message}')
            return error_message
        else:
            import_return_str = json.dumps(setup_list, ensure_ascii=False).encode('utf8').decode()
            # print(f'import_return_str:{import_return_str}')
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str

    def get_color(self, project_id):

        color_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_color_v2",
            body={
            "args": [[('job_project_id', '=', project_id)]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in color_list:
            error_message = color_list.get('error_message')
            _logger.info(f'error_message: {error_message}')
            return error_message
        else:
            import_return_str = json.dumps(color_list, ensure_ascii=False).encode('utf8').decode()
            # print(f'import_return_str:{import_return_str}')
            _logger.info(f'import_return_str: {import_return_str}')
            return import_return_str


def RegisterClass(cls):
  import os, sys, win32api, win32con
  import win32com.server.register
  file = os.path.abspath(sys.modules[cls.__module__].__file__)
  folder = os.path.dirname(file)
  module = os.path.splitext(os.path.basename(file))[0]
  python = win32com.server.register._find_localserver_exe(1)
  python = win32api.GetShortPathName(python)
  server = win32com.server.register._find_localserver_module()
  command = '%s "%s" %s' % (python, server, cls._reg_clsid_)
  typename = module + "." + cls.__name__

  def write(path, value):
    win32api.RegSetValue(win32con.HKEY_CURRENT_USER, path, win32con.REG_SZ, value)

  write("SOFTWARE\\Classes\\" + cls._reg_progid_ + '\\CLSID', cls._reg_clsid_)
  write("SOFTWARE\\Classes\\AppID\\" + cls._reg_clsid_, cls._reg_progid_)
  write("SOFTWARE\\Classes\\CLSID\\" + cls._reg_clsid_, cls._reg_desc_)
  write("SOFTWARE\\Classes\\CLSID\\" + cls._reg_clsid_ + '\\LocalServer32', command)
  write("SOFTWARE\\Classes\\CLSID\\" + cls._reg_clsid_ + '\\ProgID', cls._reg_progid_)
  write("SOFTWARE\\Classes\\CLSID\\" + cls._reg_clsid_ + '\\PythonCOMPath', folder)
  write("SOFTWARE\\Classes\\CLSID\\" + cls._reg_clsid_ + '\\PythonCOM', typename)
  write("SOFTWARE\\Classes\\CLSID\\" + cls._reg_clsid_ + '\\Debugging', "0")

  print("Registered %s" % cls.__name__)


if __name__ == '__main__':
    if '--register' in sys.argv[1:] or '--unregister' in sys.argv[1:]:
        import win32com.server.register
        win32com.server.register.UseCommandLine(ComServer)
    elif '--debug' in sys.argv[1:]:
        global debugging, useDispatcher
        debugging = 1
        from win32com.server.dispatcher import DefaultDebugDispatcher
        useDispatcher = DefaultDebugDispatcher
        import win32com.server.register
        win32com.server.register.UseCommandLine(ComServer,debug=debugging)
    elif '--user' in sys.argv[1:]:
        RegisterClass(ComServer)
    else:
        print("\n\n")
        print("     *******************************************************************************************")
        print("     *******************************************************************************************")
        print()
        print("     *** 注意：本程式為 AutoCAD 的原生程式 AutoLisp 與 Odoo 的溝通管道，請勿關閉(請縮小即可) ***")
        print()
        print("     *******************************************************************************************")
        print("     *******************************************************************************************")
        print("\n\n\n")

        # start the server.
        from win32com.server import localserver
        localserver.serve([ComServer._reg_clsid_])

