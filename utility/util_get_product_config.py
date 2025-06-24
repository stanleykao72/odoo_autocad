# -*- coding: utf-8 -*-

import json
import logging

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.INFO)

# log to console
c_handler = logging.StreamHandler()

console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
c_handler.setFormatter(console_format)
c_handler.setLevel = logging.DEBUG

_logger.addHandler(c_handler)


class GetProductConfig:
    def __init__(self, **kwargs):
        self.odoo = kwargs.get('odoo')
        self.requestOptions = kwargs.get('requestOptions')
        self.token = kwargs.get('token')

    def get_project(self, pr_no):

        project_list = self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_project_v2",
            body={
            "args": [[['name', '=', pr_no]]],
            "kwargs": {'user_token': self.user_token},
            "context": {}
            },
            _request_options=self.requestOptions, 
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
            },
            _request_options=self.requestOptions, 
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
            },
            _request_options=self.requestOptions,    
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
            },
            _request_options=self.requestOptions,  
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
