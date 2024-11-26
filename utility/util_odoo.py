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
        self.log.safe_log_insert(f"token: {token}\n")

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
