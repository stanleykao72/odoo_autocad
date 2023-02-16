#!/usr/bin/env python
# coding: utf-8

import json
import requests
import simplejson
import yaml
import csv
import re
# from pyautocad import Autocad, APoint, utils

import win32com.client
import pythoncom

from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from bravado.exception import HTTPError
from swagger_spec_validator.common import SwaggerValidationError    # type: ignore

def vtPnt(x, y, z=0):
    """座標點轉化爲浮點數"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def vtObj(obj):
    """轉化爲對象數組"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

def vtFloat(list):
    """列表轉化爲浮點數"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)

def vtInt(list):
    """列表轉化爲整數"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)

def vtVariant(list):
    """列表轉化爲變體"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list)

def iter_objects(acaddoc, object_name_or_list=None, block=None,
                 limit=None, dont_cast=False):
    """Iterate objects from `block`
    :param object_name_or_list: part of object type name, or list of it
    :param block: Autocad block, default - :class:`ActiveDocument.ActiveLayout.Block`
    :param limit: max number of objects to return, default infinite
    :param dont_cast: don't retrieve best interface for object, may speedup
                      iteration. Returned objects should be casted by caller
    """
    if block is None:
        block = acaddoc.ActiveLayout.Block
    object_names = object_name_or_list
    if object_names:
        if isinstance(object_names, str):
            object_names = [object_names]
        object_names = [n.lower() for n in object_names]

    count = block.Count
    for i in range(count):
        item = block.Item(i)  # it's faster than `for item in block`
        if limit and i >= limit:
            return
        if object_names:
            object_name = item.ObjectName.lower()
            if not any(possible_name in object_name for possible_name in object_names):
                continue
        # if not dont_cast:
        #     item = self.best_interface(item)
        yield item

def find_one(acaddoc, object_name_or_list, container=None, predicate=None):
    """Returns first occurance of object which match `predicate`
    :param object_name_or_list: like in :meth:`iter_objects`
    :param container: like in :meth:`iter_objects`
    :param predicate: callable, which accepts object as argument
                      and returns `True` or `False`
    :returns: Object if found, else `None`
    """
    if predicate is None:
        predicate = bool
    for obj in iter_objects(acaddoc, object_name_or_list, container):
        if predicate(obj):
            return obj
    return None

def prompt(acaduti, prompt_txt):
    return acaduti.Prompt(prompt_txt)

def main():

    # 連線及庫匯入
    # acad = Autocad(create_if_not_exists = True)
    # acad.prompt("Hello! Autocad from Python.")
    # print(acad.doc.Name)

    # 連線
    acad = win32com.client.Dispatch("AutoCAD.Application")
    # acad.Visible = True
    acaddoc = acad.ActiveDocument
    acaduti = acad.ActiveDocument.Utility
    prompt(acaduti, "Hello! Autocad from pywin32.\n")
    mp = acaddoc.ModelSpace
    prompt(acaduti, f'{acaddoc.Name}\n')

    file_path = acaddoc.GetVariable('dwgprefix')
    prompt(acaduti, f'{file_path}\n')

    # read config file
    with open('c:/odoo/config/token.yaml', 'r') as f:
        token_config = yaml.load(f, Loader=yaml.FullLoader)

    print(token_config)
    token_cfg = token_config['user']
    print(token_cfg)

    if 'token' in token_cfg:
        pass
    else:
        prompt(acaduti, 'token.yaml設定中沒有 token\n')
        return

    if 'server_file' in token_cfg:
        with open(f"{token_cfg['server_file']}", 'r') as f:
            server_config = yaml.load(f, Loader=yaml.FullLoader)
    
        # print(server_config)
        server_cfg = server_config['server']
        print(server_cfg)
    else:
        prompt(acaduti, 'token.yaml設定中沒有 server_file。\n請檢查c:/odoo/config/token.yaml\n')
        return

    # load autolsip
    acaddoc.SendCommand(f'(load "{server_cfg["StripMtext_file"]}")(load "{server_cfg["read_csv_file"]}")\n')

    # point0 = APoint (1,1)
    point0 = vtPnt(0, 0, 0)

    # point1 = acaduti.GetPoint ( point0, "請選擇基點")
    # point2 = acaduti.GetCorner (APoint(point1), "請選擇對角點")

    point1 = acaduti.GetPoint ( point0, "請選擇基點")
    point2 = acaduti.GetCorner (vtFloat(point1), "請選擇對角點")
    prompt(acaduti, f'{point1}, {point2}\n')

    # try:
    #     ssget1 = acaddoc.SelectionSets.Add("SS1")
    # except:
    #     acaddoc.SelectionSets.Item("SS1").Delete
    #     ssget1 = acaddoc.SelectionSets("SS1")
    #     ssget1.Clear()

    # ssget1.select(1, APoint(point1), APoint(point2))

    try:
        ssget = acaddoc.SelectionSets.Add("SS")
    except:
        acaddoc.SelectionSets.Item("SS").Delete
        ssget = acaddoc.SelectionSets("SS")
        ssget.Clear()

    pnts = point1 + point2
    # print(pnts)
    pnts=vtFloat(pnts)
    ssget.SelectByPolygon(2, pnts)

    blocks = acaddoc.Blocks
    # blocks = acad.best_interface(blocks)

    header_dict = {}
    for block in blocks:
        if block.name == 'pr_no':
            prompt(acaduti, f'{block.name}\n')
            # print(block.ObjectName)
            item = find_one(acaddoc, "text", block)
            header_dict['pr_no'] = item.TextString
            prompt(acaduti, f'{item.TextString}\n')
        if block.name == 'prj_name':
            prompt(acaduti, f'{block.name}\n')
            # print(block.ObjectName)
            item = find_one(acaddoc, "text", block)
            header_dict['project_name'] = item.TextString
            prompt(acaduti, f'{item.TextString}\n')

    host = server_cfg['host']  #'odoo-esmith-1124-stage-6571675.dev.odoo.com'
    db_name = server_cfg['db_name'] #'odoo-esmith-1124-stage-6571675'
    user_token = token_cfg['token'] #'6d4bead3-8c1a-46b4-a399-7e57535b85d9'
    url = server_cfg['url'] #'https://odoo-esmith-1124-stage-6571675.dev.odoo.com/api/v1/boq_import_api/swagger.json?token=34dba8ba-cf29-4ac7-a2b7-e64b7ac7bae6&db=odoo-esmith-1124-stage-6571675'
    http_client = RequestsClient()
    http_client.set_basic_auth(host, db_name, user_token)

    odoo = False
    try:
        odoo = SwaggerClient.from_url(
            url,
            http_client=http_client
        )
        prompt(acaduti, f"與Odoo連線成功\n")
    except requests.exceptions.ConnectionError:
        prompt(acaduti, 'Unable to connect to server.\n')
        prompt(acaduti, f"無法與Odoo，通常多試幾次會成功\n")
    except (
        simplejson.errors.JSONDecodeError,      # type: ignore
        yaml.YAMLError,
        HTTPError,
        ):
        print(
            'Invalid swagger file. Please check to make sure the '
            'swagger file can be found at: {}.'.format(url)
        )
    except SwaggerValidationError:
        print('Invalid swagger format.')

    if odoo:
        product_list = odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_product",
            body={
            "args": [[('categ_id', 'child_of', 27), ('active', '=', True)]],
            "kwargs": {'user_token': user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in product_list:
            print(product_list)
            prompt(acaduti, f"錯誤:{product_list.get('error_code')} ==> {product_list.get('error_message')}")
        else:
            product_file = f'{file_path}product.csv'
            with open(product_file, 'w', encoding='utf-8', newline='') as csvfile:
                # 建立 CSV 檔寫入器
                # writer = csv.writer(csvfile)
                fieldnames = ['id', 'name', 'uom']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                # writer.writeheader()
                for product in product_list:
                    writer.writerow(product)

            product_file = product_file.replace(chr(92), chr(92)+chr(92))
            acaddoc.SendCommand(f'(utf_convert "{product_file}" "utf-8" "{product_file}" "ANSI")\n')

        prompt(acaduti, f"product convert to ansi format \n")

        # contract_list = odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
        #     method_name="get_contract",
        #     body={
        #     "args": [[('name', '=', header_dict['pr_no'])]],
        #     "kwargs": {'user_token': user_token, 'prj_name': header_dict['project_name']},
        #     "context": {}
        #     }    
        # ).response().incoming_response.json()

        # if 'error_code' in contract_list:
        #     print(contract_list)
        #     prompt(acaduti, f"錯誤:{contract_list.get('error_code')} ==> {contract_list.get('error_message')}")
        # else:
        #     contract_file = f'{file_path}contract.csv'
        #     with open(contract_file, 'w', encoding='utf-8', newline='') as csvfile:
        #         # 建立 CSV 檔寫入器
        #         # writer = csv.writer(csvfile)
        #         fieldnames = ['id', 'reference', 'contract_item']
        #         writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        #         # writer.writeheader()
        #         for contract in contract_list:
        #             writer.writerow(contract)

        #     contract_file = contract_file.replace(chr(92), chr(92)+chr(92))
        #     acaddoc.SendCommand(f'(utf_convert "{contract_file}" "utf-8" "{contract_file}" "ANSI")\n')    

        setup_list = odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_setup",
            body={
            "args": [[]],
            "kwargs": {'user_token': user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in setup_list:
            print(setup_list)
            prompt(acaduti, f"錯誤:{setup_list.get('error_code')} ==> {setup_list.get('error_message')}\n")
        else:
            setup_file = f'{file_path}setup.csv'
            with open(setup_file, 'w', encoding='utf-8', newline='') as csvfile:
                # 建立 CSV 檔寫入器
                # writer = csv.writer(csvfile)
                fieldnames = ['key', 'value']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                # writer.writeheader()
                for setup in setup_list:
                    writer.writerow(setup)

            setup_file = setup_file.replace(chr(92), chr(92)+chr(92))
            acaddoc.SendCommand(f'(utf_convert "{setup_file}" "utf-8" "{setup_file}" "ANSI")\n')    

        project_list = odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_project",
            body={
            "args": [[('name', '=', header_dict['pr_no'])]],
            "kwargs": {'user_token': user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if 'error_code' in project_list:
            print(project_list)
            prompt(acaduti, f"錯誤:{project_list.get('error_code')} ==> {project_list.get('error_message')}\n")

        color_list = odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
            method_name="get_color",
            body={
            "args": [[('job_project_id', '=', project_list[0].get('id'))]],
            "kwargs": {'user_token': user_token},
            "context": {}
            }    
        ).response().incoming_response.json()

        if color_list and 'error_code' in color_list:
            print(setup_list)
            prompt(acaduti, f"錯誤:{color_list.get('error_code')} ==> {color_list.get('error_message')}")
        else:
            color_file = f'{file_path}color.csv'
            with open(color_file, 'w', encoding='utf-8', newline='') as csvfile:
                # 建立 CSV 檔寫入器
                # writer = csv.writer(csvfile)
                fieldnames = ['name', 'color_no']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                # writer.writeheader()
                for color in color_list:
                    writer.writerow(color)

            color_file = color_file.replace(chr(92), chr(92)+chr(92))
            acaddoc.SendCommand(f'(utf_convert "{color_file}" "utf-8" "{color_file}" "ANSI")\n')

        prompt(acaduti, f"已在 {file_path} 產生csv檔\n")

    for entity in ssget:
        name = entity.EntityName
        print(f'Name:{name}\n')
        if name == 'AcDbBlockReference':
            print(f'name:{entity.Name}\n')
            print(f'{entity.EffectiveName}')
            for attrib in entity.GetAttributes():
                if attrib.TagString == 'job_working_plan_name':
                    attrib.MTextAttribute = True
                    print(f'update 工種群組 ')
                    attrib.TextString = project_list[0].get('job_working_plan_name')
                    attrib.UpdateMTextAttribute()
                if attrib.TagString == 'project_name':
                    attrib.MTextAttribute = True
                    print(f'update 工程名稱 ')
                    attrib.TextString = project_list[0].get('name')
                    attrib.UpdateMTextAttribute()
                attrib.Update()

    prompt(acaduti, f"已更新圖框中的工程名稱 {project_list[0].get('name')} 及工種群組 {project_list[0].get('job_working_plan_name')} \n")

if __name__ == "__main__":
    main()
