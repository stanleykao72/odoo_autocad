# -*- coding: utf-8 -*-
import logging
import base64
import requests
import tkinter as tk
from tkinter import ttk, messagebox

from utility.util_odoo import UtilOdoo
from utility.util_autocad import UtilAutoCAD
from utility.util_push_to_boq import UtilPushToBoq
from utility.util_transfer_boq_to_pr import UtilTransferBoqToPr
from forms.form_autocad_param import FormAutoCADParam
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.INFO)

# log to console
c_handler = logging.StreamHandler()

console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
c_handler.setFormatter(console_format)
c_handler.setLevel = logging.DEBUG

_logger.addHandler(c_handler)

class FormMain(tk.Tk):
    def __init__(self, odoo=None, requestOptions=None, token=None, odoo_connection=None):
        super().__init__()
        self.odoo = odoo
        self.requestOptions = requestOptions
        self.token = token
        self.odoo_connection = odoo_connection
        self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
        self.autocad_util = UtilAutoCAD()
        self.push_to_boq_util = UtilPushToBoq()
        self.transfer_boq_to_pr_util = UtilTransferBoqToPr()
        self.filtered_items = []  # 初始化 filtered_items
        
        # 設置視窗標題和大小
        self.title("材料及加工方式")
        window_width = 700
        window_height = 450
        
        # 獲取螢幕尺寸
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 計算視窗位置 (靠右)
        x = screen_width - window_width
        y = (screen_height - window_height) // 2  # 垂直置中
        
        # 設置視窗位置和大小
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.create_widgets()

    def create_widgets(self):
        self.top_bar = tk.Frame(self, bg="orange", height=50)
        self.top_bar.pack(side="top", fill="x")

        self.side_menu = tk.Frame(self, bg="lightgrey", width=200)
        self.side_menu.pack(side="left", fill="y")

        self.main_body = tk.Frame(self, bg="white")
        self.main_body.pack(side="right", fill="both", expand=True)

        self.create_side_menu()

    def create_side_menu(self):
        self.menu_button_odoo = tk.Button(self.side_menu, text="連接到 Odoo", command=self.connect_odoo)
        self.menu_button_odoo.pack(fill="x", pady=5)

        self.menu_button_autocad = tk.Button(self.side_menu, text="連接到 AutoCAD", command=self.connect_autocad)
        self.menu_button_autocad.pack(fill="x", pady=5)

        self.menu_button_get_params = tk.Button(self.side_menu, text="從 Odoo 獲取參數", command=self.get_parameters_from_odoo)
        self.menu_button_get_params.pack(fill="x", pady=5)

        self.menu_button_push_boq = tk.Button(self.side_menu, text="推送到 BOQ", command=self.push_to_boq)
        self.menu_button_push_boq.pack(fill="x", pady=5)

        self.menu_button_transfer_boq = tk.Button(self.side_menu, text="轉移 BOQ 到 PR", command=self.transfer_boq_to_pr)
        self.menu_button_transfer_boq.pack(fill="x", pady=5)

    def connect_odoo(self):
        self.clear_main_body()
        self.log_messages = tk.Text(self.main_body)
        self.log_messages.pack(fill="both", expand=True)
        self.log_messages.insert(tk.END, "連接到 Odoo...\n")
        try:
            odoo, requestOptions, token = self.odoo_util.connect_odoo(self.odoo_connection)
            self.after(1000, lambda: self.log_messages.insert(tk.END, "與 Odoo 連線成功\n"))
            self.odoo = odoo
            self.requestOptions = requestOptions
            self.token = token
            self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
        except requests.exceptions.ConnectionError:
            self.after(1000, lambda: self.log_messages.insert(tk.END, "無法與 Odoo 連線，通常多試幾次會成功\n"))
        except Exception as e:
            self.after(1000, lambda: self.log_messages.insert(tk.END, f"連接 Odoo 時發生錯誤: {str(e)}\n"))

    def connect_autocad(self):
        self.autocad_util.connect_autocad(self.main_body)

    def get_parameters_from_odoo(self):
        form_autocad_param = FormAutoCADParam(self.main_body, self.odoo_util)
        form_autocad_param.get_parameters_from_odoo()

    def push_to_boq(self):
        self.push_to_boq_util.push_to_boq(self.main_body)

    def transfer_boq_to_pr(self):
        self.transfer_boq_to_pr_util.transfer_boq_to_pr(self.main_body)

    def clear_main_body(self):
        for widget in self.main_body.winfo_children():
            widget.destroy()

if __name__ == '__main__':
    odoo_connection = {
        'host': 'your_host',
        'db_name': 'your_db_name',
        'url': 'your_url',
        'token': 'your_token'
    }
    app = FormMain(odoo_connection=odoo_connection)
    app.mainloop()
