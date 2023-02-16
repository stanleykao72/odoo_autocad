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
    acaddoc.SendCommand(f'(load "{server_cfg["contract_product_file"]}")\n')

    product_file = f'{file_path}product.csv'
    # contract_file = f'{file_path}contract.csv'
    setup_file = f'{file_path}setup.csv'
    color_file = f'{file_path}color.csv'

    product_file = product_file.replace(chr(92), chr(92)+chr(92))
    # contract_file = contract_file.replace(chr(92), chr(92)+chr(92))
    setup_file = setup_file.replace(chr(92), chr(92)+chr(92))
    color_file = color_file.replace(chr(92), chr(92)+chr(92))

    acaddoc.SendCommand(f'(contract_product "{product_file}" "{setup_file}" "{color_file}")\n')

if __name__ == "__main__":
    main()
