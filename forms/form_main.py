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
from utility.util_log import UtilLog
from forms.form_autocad_param import FormAutoCADParam
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

class FormMain(tk.Tk):
    def __init__(self, odoo_connection):
        super().__init__()
        self.odoo_connection = odoo_connection
        self.title("AutoCAD Odoo Integration")

        # 獲取屏幕寬度
        screen_width = self.winfo_screenwidth()
        window_width = 800
        window_height = 600

        # 計算窗口應該放置的位置
        x_position = screen_width - window_width
        y_position = 0

        # 設置窗口大小並將其放置在屏幕的最右邊
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Define main layout frames
        self.top_bar = tk.Frame(self, bg="orange", height=50)
        self.top_bar.pack(side="top", fill="x")

        self.side_menu = tk.Frame(self, width=200, bg="lightgray")
        self.side_menu.pack(side="left", fill="y")

        self.main_content = tk.Frame(self, bg="white")
        self.main_content.pack(side="top", fill="both", expand=True)

        self.log_frame = tk.Frame(self, bg="black")
        self.log_frame.pack(side="bottom", fill="x")

        # Initialize log utility
        self.log_util = UtilLog(self.log_frame)

        # Initialize Odoo and AutoCAD utilities
        self.odoo_util = UtilOdoo(self.odoo_connection, self.log_util)
        self.autocad_util = UtilAutoCAD(self.odoo_util, self.log_util)
        self.push_to_boq_util = UtilPushToBoq(self.odoo_util, self.autocad_util, self.log_util)
        self.transfer_boq_to_pr_util = UtilTransferBoqToPr(self.odoo_util, self.autocad_util, self.log_util)

        # Add buttons to the side menu
        self.add_side_menu_buttons()

    def add_side_menu_buttons(self):
        buttons = [
            ("連接到 Odoo", self.connect_odoo, 'button_connect_odoo'),
            ("連接到 AutoCAD", self.connect_autocad, 'button_connect_autocad'),
            ("從 Odoo 獲取參數", self.get_parameters_from_odoo, 'button_get_parameters'),
            ("推送到 BOQ", self.push_to_boq, 'button_push_to_boq'),
            ("轉移 BOQ 到 PR", self.transfer_boq_to_pr, 'button_transfer_boq_to_pr'),
            ("清除此配置表格ID", self.clear_table_id, 'button_clear_table_id'),
            ("清除所有配置表格ID", self.clear_all_tables_id, 'button_clear_all_tables_id'),
        ]

        for (text, command, button_name) in buttons:
            button = tk.Button(self.side_menu, text=text, command=command, name=button_name)
            button.pack(fill="x", pady=5)
            setattr(self, button_name, button)  # 將按鈕賦值為類的屬性


        if self.odoo_util.connected_odoo():
            self.button_connect_odoo.config(bg="green")
        else:
            self.button_connect_odoo.config(bg="red")

        if self.autocad_util.connected_autocad():
            self.button_connect_autocad.config(bg="green")
        else:
            self.button_connect_autocad.config(bg="red")

    def connect_odoo(self):
        self.clear_main_content()
        self.log_util.safe_log_insert("連接到 Odoo...\n")
        try:
            odoo, requestOptions, token = self.odoo_util.connect_odoo(self.odoo_connection)
            # self.after(1000, lambda: self.log_util.safe_log_insert("與 Odoo 連線成功\n"))
            if self.odoo_util.connected_odoo():
                self.button_connect_odoo.config(bg="green")
        except requests.exceptions.ConnectionError:
            self.after(1000, lambda: self.log_util.safe_log_insert("無法與 Odoo 連線，通常多試幾次會成功\n"))
        except Exception as e:
            self.after(1000, lambda: self.log_util.safe_log_insert(f"連接 Odoo 時發生錯誤: {str(e)}\n"))

    def connect_autocad(self):
        self.autocad_util.connect_autocad(self.main_content)
        if self.autocad_util.connected_autocad():
            self.button_connect_autocad.config(bg="green")

    def get_parameters_from_odoo(self):
        # Ensure AutoCAD is connected and attributes are available
        if not self.autocad_util.project_id:
            self.show_error_message("錯誤", "請先連接 AutoCAD 並確保已獲取專案資料。")
            return

        # Initialize FormAutoCADParam with the UtilAutoCAD instance
        form_autocad_param = FormAutoCADParam(
            main_content=self.main_content,
            odoo_util=self.odoo_util,
            autocad_util=self.autocad_util,
            log_util=self.log_util,
            root=self
        )
        form_autocad_param.get_parameters_from_odoo()

    def push_to_boq(self):
        # Implement push to BOQ logic
        self.push_to_boq_util.push_to_boq()

    def transfer_boq_to_pr(self):
        # Implement transfer BOQ to PR logic
        self.transfer_boq_to_pr_util.transfer_boq_to_pr()
    
    def clear_table_id(self):
        self.autocad_util.clear_table_id()

    def clear_all_tables_id(self):
        self.autocad_util.clear_all_tables_id()

    def clear_main_content(self):
        for widget in self.main_content.winfo_children():
            widget.destroy()

    def show_error_message(self, title, message):
        self.update_idletasks()  # 確保窗口已更新
        messagebox.showerror(title, message, parent=self)

    def show_info_message(self, title, message):
        self.update_idletasks()  # 確保窗口已更新
        messagebox.showinfo(title, message, parent=self)

# class FormMain(tk.Tk):
#     def __init__(self, odoo=None, requestOptions=None, token=None, odoo_connection=None):
#         super().__init__()
#         self.odoo = odoo
#         self.requestOptions = requestOptions
#         self.token = token
#         self.odoo_connection = odoo_connection

#         self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
#         self.autocad_util = UtilAutoCAD(self.odoo_util)
#         self.push_to_boq_util = UtilPushToBoq()
#         self.transfer_boq_to_pr_util = UtilTransferBoqToPr()
#         self.filtered_items = []  # 初始化 filtered_items
        
#         # 設置視窗標題和大小
#         self.title("材料及加工方式")
#         window_width = 700
#         window_height = 450
        
#         # 獲取螢幕尺寸
#         screen_width = self.winfo_screenwidth()
#         screen_height = self.winfo_screenheight()
        
#         # 計算視窗位置 (靠右)
#         x = screen_width - window_width
#         y = (screen_height - window_height) // 2  # 垂直置中
        
#         # 設置視窗位置和大小
#         self.geometry(f"{window_width}x{window_height}+{x}+{y}")

#         # 初始化日誌訊息
#         self.autocad_util.initialize_log(self.log_frame)

#         self.create_widgets()

#     def create_widgets(self):
#         self.top_bar = tk.Frame(self, bg="orange", height=50)
#         self.top_bar.pack(side="top", fill="x")

#         self.side_menu = tk.Frame(self, bg="lightgrey", width=200)
#         self.side_menu.pack(side="left", fill="y")

#         self.main_body = tk.Frame(self, bg="white")
#         self.main_body.pack(side="right", fill="both", expand=True)

#         self.create_side_menu()

#     def create_side_menu(self):
#         self.menu_button_odoo = tk.Button(self.side_menu, text="連接到 Odoo", command=self.connect_odoo)
#         self.menu_button_odoo.pack(fill="x", pady=5)

#         self.menu_button_autocad = tk.Button(self.side_menu, text="連接到 AutoCAD", command=self.connect_autocad)
#         self.menu_button_autocad.pack(fill="x", pady=5)

#         self.menu_button_get_params = tk.Button(self.side_menu, text="從 Odoo 獲取參數", command=self.get_parameters_from_odoo)
#         self.menu_button_get_params.pack(fill="x", pady=5)

#         self.menu_button_push_boq = tk.Button(self.side_menu, text="推送到 BOQ", command=self.push_to_boq)
#         self.menu_button_push_boq.pack(fill="x", pady=5)

#         self.menu_button_transfer_boq = tk.Button(self.side_menu, text="轉移 BOQ 到 PR", command=self.transfer_boq_to_pr)
#         self.menu_button_transfer_boq.pack(fill="x", pady=5)

#     def connect_odoo(self):
#         self.clear_main_body()
#         self.log_messages = tk.Text(self.main_body)
#         self.log_messages.pack(fill="both", expand=True)
#         self.log_messages.insert(tk.END, "連接到 Odoo...\n")
#         try:
#             self.log_messages.insert(tk.END, f"self.odoo_connection: {self.odoo_connection}\n")
#             odoo, requestOptions, token = self.odoo_util.connect_odoo(self.odoo_connection)

#             log = self.log_messages
#             self.after(1000, lambda: log.insert(tk.END, "與 Odoo 連線成功\n"))
#             self.odoo = odoo
#             self.requestOptions = requestOptions
#             self.token = token
#             self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
#         except requests.exceptions.ConnectionError:
#             log = self.log_messages
#             self.after(1000, lambda: log.insert(tk.END, "無法與 Odoo 連線，通常多試幾次會成功\n"))
#         except Exception as e:
#             log = self.log_messages
#             self.after(1000, lambda: log.insert(tk.END, f"連接 Odoo 時發生錯誤: {str(e)}\n"))

#     def connect_autocad(self):
#         self.autocad_util.connect_autocad(self.main_body)

#     def get_parameters_from_odoo(self):
#         # Ensure AutoCAD is connected and attributes are available
#         if not self.autocad_util.project_id:
#             messagebox.showerror("錯誤", "請先連接 AutoCAD 並確保已獲取專案資料。")
#             return

#         # Extract required attributes
#         project_id = self.autocad_util.project_id

#         # Initialize FormAutoCADParam with extracted attributes
#         form_autocad_param = FormAutoCADParam(
#             main_body=self.main_body,
#             odoo_util=self.odoo_util,
#             util_autocad=self.autocad_util  # Pass the UtilAutoCAD instance
#         )
#         form_autocad_param.get_parameters_from_odoo()

#     def push_to_boq(self):
#         self.push_to_boq_util.push_to_boq(self.main_body)

#     def transfer_boq_to_pr(self):
#         self.transfer_boq_to_pr_util.transfer_boq_to_pr(self.main_body)

#     def clear_main_body(self):
#         for widget in self.main_body.winfo_children():
#             widget.destroy()

if __name__ == '__main__':
    odoo_connection = {
        'host': 'your_host',
        'db_name': 'your_db_name',
        'url': 'your_url',
        'token': 'your_token'
    }
    app = FormMain(odoo_connection=odoo_connection)
    app.mainloop()
