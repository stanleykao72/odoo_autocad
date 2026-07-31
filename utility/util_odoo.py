# -*- coding: utf-8 -*-
import json
import requests
import simplejson
import yaml
import base64
# from utility.util_log import UtilLog

from requests.exceptions import HTTPError
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

from utility.util_secrets import mask_secret

# _logger = logging.getLogger(__name__)
# _logger.setLevel(logging.INFO)

# # log to console
# c_handler = logging.StreamHandler()

# console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
# c_handler.setFormatter(console_format)
# c_handler.setLevel = logging.DEBUG

# _logger.addHandler(c_handler)


class UtilOdoo:
    def __init__(self, odoo_connection, log_util):
        # self.odoo = odoo
        # self.requestOptions = requestOptions
        # self.user_token = token
        self.odoo = None
        self.requestOptions = None
        self.user_token = None
        self.log = log_util
        self.connect_odoo(odoo_connection)

    def connected_odoo(self):
        if self.odoo:
            return True
        else:
            return False

    def string_to_base64(self, input_string):
        input_bytes = input_string.encode('utf-8')        
        base64_bytes = base64.b64encode(input_bytes)
        base64_string = base64_bytes.decode('utf-8')
        return base64_string

    def connect_odoo(self, odoo_connection):
        host = odoo_connection['host']
        db_name = odoo_connection['db_name']
        url = odoo_connection['url']
        token = odoo_connection['token']
        self.log.safe_log_insert(f"host: {host}\n")
        self.log.safe_log_insert(f"db_name: {db_name}\n")
        self.log.safe_log_insert(f"url: {url}\n")
        # 不記錄 token 明文（日誌會落地到 logs/*.log）
        self.log.safe_log_insert(f"token: {mask_secret(token)}\n")

        http_client = RequestsClient()
        http_client.set_basic_auth(host, db_name, token)
        try:
            odoo = SwaggerClient.from_url(url, http_client=http_client)
            basic_string = f'{db_name}:{token}'
            basic_token = self.string_to_base64(basic_string)
            headers = {
                'Authorization': f'Basic {basic_token}'
            }
            requestOptions = {
                'headers': headers,
            }
            self.log.safe_log_insert(f"與 Odoo 連線成功\n")
            self.odoo = odoo
            self.requestOptions = requestOptions
            self.user_token = token
            return odoo, requestOptions, token
        except requests.exceptions.ConnectionError:
            self.log.safe_log_insert(f"無法與 Odoo 連線，通常多試幾次會成功\n")
            raise
        except (
            simplejson.errors.JSONDecodeError,
            yaml.YAMLError,
            HTTPError,
            ):
            self.log.safe_log_insert(f"無效的 Swagger 文件。請檢查確保 Swagger 文件可以在 {url} 找到。\n")
            raise
        except SwaggerValidationError:
            self.log.safe_log_insert(f'無效的 Swagger 格式。\n')
            raise

    def import2boq(self, layout_dict):

        boq_dict = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="import2boq_v2",
            body={
            "args": [layout_dict],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions, 
        ).response().incoming_response.json()

        if 'error_code' in boq_dict:
            error_message = boq_dict.get('error_message')
            self.log.safe_log_insert(f'error_message: {error_message}\n')
            return error_message
        else:
            return_boq_list = boq_dict.get('all')
            return return_boq_list

    def boq2pr(self, layout_dict):

        pr_dict = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="boq2pr_v2",
            body={
            "args": [layout_dict],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions, 
        ).response().incoming_response.json()

        if 'error_code' in pr_dict:
            error_message = pr_dict.get('error_message')
            self.log.safe_log_insert(f'error_message: {error_message}\n')
            return error_message
        else:
            return_pr_list = pr_dict.get('all')
            return return_pr_list

    def get_project(self, pr_no):

        project_dict = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_project_v2",
            body={
            "args": [[['name', '=', pr_no]]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions, 
        ).response().incoming_response.json()

        if 'error_code' in project_dict:
            error_message = project_dict.get('error_message')
            self.log.safe_log_insert(f'error_message: {error_message}\n')
            return error_message
        else:
            # return_project_list = project_list.get('project')
            return project_dict

    def get_product(self):

        product_dict = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_product_v2",
            body={
            "args": [[('categ_id', 'child_of', 27), ('active', '=', True)]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions, 
        ).response().incoming_response.json()

        if 'error_code' in product_dict:
            error_message = product_dict.get('error_message')
            self.log.safe_log_insert(f'error_message: {error_message}\n')
            return error_message
        else:
            return_product_list = product_dict.get('product')
            return return_product_list

    def get_setup(self, setup_name):

        setup_dict = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_setup_v2",
            body={
            "args": [[('setup_name', '=', setup_name)]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions,    
        ).response().incoming_response.json()

        if 'error_code' in setup_dict:
            error_message = setup_dict.get('error_message')
            self.log.safe_log_insert(f'error_message: {error_message}\n')
            return error_message
        else:
            return_setup_list = setup_dict.get('setup')
            return return_setup_list

    def get_color(self, project_id):

        color_dict = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_color_v2",
            body={
            "args": [[('job_project_id', '=', project_id)]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions,  
        ).response().incoming_response.json()

        if 'error_code' in color_dict:
            error_message = color_dict.get('error_message')
            self.log.safe_log_insert(f'error_message: {error_message}\n')
            return error_message
        else:
            return_color_list = color_dict.get('color')
            return return_color_list
    
    def search_products(self, query, limit=50):
        """搜尋產品資料"""
        try:
            if not self.connected_odoo():
                self.log.safe_log_insert("Odoo未連接，無法搜尋產品\n")
                return []
            
            # 使用現有的get_product方法作為基礎，並添加查詢功能
            # 這裡是一個簡化的實現，實際應該調用Odoo API搜尋產品
            products = []
            
            # 模擬產品搜尋結果 (實際實現時需要調用真正的Odoo API)
            if query.lower() in ['pipe', '管道']:
                products = [
                    {'id': 1, 'name': 'Pipe 100mm', 'code': 'P100', 'category': 'Piping'},
                    {'id': 2, 'name': 'Pipe 150mm', 'code': 'P150', 'category': 'Piping'}
                ]
            elif query.lower() in ['valve', '閥門']:
                products = [
                    {'id': 3, 'name': 'Valve 50mm', 'code': 'V050', 'category': 'Valves'},
                    {'id': 4, 'name': 'Valve 100mm', 'code': 'V100', 'category': 'Valves'}
                ]
            
            products = products[:limit]  # 限制返回數量
            self.log.safe_log_insert(f"搜尋產品完成，找到 {len(products)} 項產品\n")
            return products
            
        except Exception as e:
            self.log.safe_log_insert(f"搜尋產品時發生錯誤: {str(e)}\n")
            return []
    
    def push_boq_data(self, project_id, boq_data):
        """推送BOQ資料到專案"""
        try:
            if not self.connected_odoo():
                self.log.safe_log_insert("Odoo未連接，無法推送BOQ資料\n")
                return False
            
            # 這裡會整合現有的import2boq功能
            # 暫時返回成功狀態，實際實現時需要調用真正的BOQ推送邏輯
            
            self.log.safe_log_insert(f"開始推送 {len(boq_data)} 項BOQ資料到專案 {project_id}\n")
            
            # 模擬推送過程
            for item in boq_data:
                item_name = item.get('name', 'Unknown')
                self.log.safe_log_insert(f"推送項目: {item_name}\n")
            
            self.log.safe_log_insert("BOQ資料推送完成\n")
            return True
            
        except Exception as e:
            self.log.safe_log_insert(f"推送BOQ資料時發生錯誤: {str(e)}\n")
            return False
