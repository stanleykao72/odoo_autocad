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

def unformat_mtext(s, exclude_list=('P', 'S')):
    """Returns string with removed format information
    :param s: string with multitext
    :param exclude_list: don't touch tags from this list. Default ('P', 'S') for
                         newline and fractions
    ::
        >>> text = ur'{\\fGOST type A|b0|i0|c204|p34;TEST\\fGOST type A|b0|i0|c0|p34;123}'
        >>> unformat_mtext(text)
        u'TEST123'
    """
    s = re.sub(r'\{?\\[^%s][^;]+;' % ''.join(exclude_list), '', s)
    s = re.sub(r'\}', '', s)
    return s


def mtext_to_string(s):
    """
    Returns string with removed format innformation as :func:`unformat_mtext` and
    `\\P` (paragraphs) replaced with newlines
    ::
        >>> text = ur'{\\fGOST type A|b0|i0|c204|p34;TEST\\fGOST type A|b0|i0|c0|p34;123}\\Ptest321'
        >>> mtext_to_string(text)
        u'TEST123\\ntest321'
    """

    return unformat_mtext(s).replace(u'\\P', u'\n')


def string_to_mtext(s):
    """Returns string in Autocad multitext format
    Replaces newllines `\\\\n` with `\\\\P`, etc.
    """
    return s.replace('\\', '\\\\').replace(u'\n', u'\P')

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
    acadlyt = acad.ActiveDocument.Layouts
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

    # point0 = APoint (1,1)
    # point0 = vtPnt(0, 0, 0)
    # point1 = acaduti.GetPoint ( point0, "請選擇基點")
    # point2 = acaduti.GetCorner (APoint(point1), "請選擇對角點")

    # point1 = acaduti.GetPoint ( point0, "請選擇基點")
    # point2 = acaduti.GetCorner (vtFloat(point1), "請選擇對角點")
    # prompt(acaduti, f'{point1}, {point2}\n')

    # try:
    #     ssget1 = acaddoc.SelectionSets.Add("SS1")
    # except:
    #     acaddoc.SelectionSets.Item("SS1").Delete
    #     ssget1 = acaddoc.SelectionSets("SS1")
    #     ssget1.Clear()

    # ssget1.select(1, APoint(point1), APoint(point2))

    # try:
    #     ssget = acaddoc.SelectionSets.Add("SS")
    # except:
    #     acaddoc.SelectionSets.Item("SS").Delete
    #     ssget = acaddoc.SelectionSets("SS")
    #     ssget.Clear()
        
    # pnts = point1 + point2
    # print(pnts)
    # pnts=vtFloat(pnts)
    # ssget.SelectByPolygon(2, pnts)

    host = server_cfg['host']  #'odoo-esmith-1124-stage-6571675.dev.odoo.com'
    db_name = server_cfg['db_name'] #'odoo-esmith-1124-stage-6571675'
    user_token = token_cfg['token'] #'6d4bead3-8c1a-46b4-a399-7e57535b85d9'
    url = server_cfg['url'] #'https://odoo-esmith-1124-stage-6571675.dev.odoo.com/api/v1/boq_import_api/swagger.json?token=34dba8ba-cf29-4ac7-a2b7-e64b7ac7bae6&db=odoo-esmith-1124-stage-6571675'
    http_client = RequestsClient()
    http_client.set_basic_auth(host, db_name, user_token)

    try:
        odoo = SwaggerClient.from_url(
            url,
            http_client=http_client
        )
        prompt(acaduti, f"與Odoo連線成功\n")
    except requests.exceptions.ConnectionError:
        print('Unable to connect to server.')
        prompt(acaduti, f"無法與Odoo，通常多試幾次會成功\n")
        return
    except (
        simplejson.errors.JSONDecodeError,      # type: ignore
        yaml.YAMLError,
        HTTPError,
        ):
        print(
            'Invalid swagger file. Please check to make sure the '
            'swagger file can be found at: {}.'.format(url)
        )
        return
    except SwaggerValidationError:
        print('Invalid swagger format.')
        return

    for lyt in acadlyt:
        print(lyt.Name)
        prompt(acaduti, f'配置: {lyt.Name}\n')
        # for block in blocks:
        header_dict = {}
        if lyt.Name != 'Model':
            table_count = 0
            detail_lst = []

            for entity in lyt.Block:
                name = entity.EntityName

                if name == 'AcDbTable':
                    table = entity
           
                    detail_flag = False
                    prompt(acaduti, f'columns: {table.Columns}, rows: {table.Rows}')
                    chk_cell_value = mtext_to_string(table.GetText(0, 0))
                    header_id = mtext_to_string(table.GetText(0, 8))
                    header_dict['header_id'] = header_id

                    if chk_cell_value == '加工細節':
                        print(f'Name:{name}\n')
                        if not 'header_id' in header_dict:
                            prompt(acaduti, f'{lyt.Name}:尚未匯入BOQ，請確認\n')
                        else:    
                            prompt(acaduti, f'{lyt.Name}:header_dict:{header_dict}\n')

                            import_return_list = odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
                                method_name="boq2pr",
                                body={
                                "args": [header_dict],
                                "kwargs": {'user_token': user_token},
                                "context": {}
                                }    
                            ).response().incoming_response.json()

                            if 'error_code' in import_return_list:
                                print(import_return_list)
                                prompt(acaduti, f"錯誤:{import_return_list.get('error_code')} ==> {import_return_list.get('error_message')}\n")
                            else:
                                hd_id = import_return_list.get('header_id')
                                hd_state = import_return_list.get('state')

                                if hd_state == 'done':
                                    prompt(acaduti, f"成功匯入請購單\n")

if __name__ == "__main__":
    main()