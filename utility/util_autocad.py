# -*- coding: utf-8 -*-
import tkinter as tk
from win32com import client
import pythoncom

from utility.util_odoo import UtilOdoo

class UtilAutoCAD:
    def __init__(self, odoo_util):
        self.acad = None
        self.doc = None
        self.pr_no = None
        self.project_id = None
        self.project_name = None
        self.job_working_plan_id = None
        self.job_working_plan_name = None
        self.log_messages = None
        self.odoo_util = odoo_util

    def initialize_log(self, log_frame):
        self.log_messages = tk.Text(log_frame, height=10, bg="black", fg="white")
        self.log_messages.pack(fill="both", expand=True)
        self.log_messages.insert(tk.END, "Log Initialized.\n")

    def connect_autocad(self, main_body):
        self.clear_main_body(main_body)
        self.log_messages = tk.Text(main_body)
        self.log_messages.pack(fill="both", expand=True)
        self.log_messages.insert(tk.END, "連接到 AutoCAD...\n")
        
        try:
            # 初始化 COM 物件
            pythoncom.CoInitialize()
            # 獲取 AutoCAD 應用程序
            self.acad = client.Dispatch("AutoCAD.Application")
            # 獲取當前文檔
            self.doc = self.acad.ActiveDocument
            
            if self.acad and self.doc:
                # 獲取檔案名稱及路徑
                file_path = self.doc.FullName
                file_name = self.doc.Name
                self.log_messages.insert(tk.END, f"AutoCAD 檔案名稱: {file_name}\n")
                self.log_messages.insert(tk.END, f"AutoCAD 檔案路徑: {file_path}\n")
                
                # 獲取並處理 PR No 及相關專案資料
                self.process_pr_no()
                
                self.log_messages.insert(tk.END, "成功連接到 AutoCAD。\n")
            else:
                self.log_messages.insert(tk.END, "無法取得 AutoCAD 文件，請確認 AutoCAD 已開啟且有文件。\n")
                
        except Exception as e:
            main_body.after(1000, lambda e=e: self.log_messages.insert(tk.END, f"連接 AutoCAD 時發生錯誤: {str(e)}\n"))

    def safe_log_insert(self, message):
        """安全地插入日誌訊息，避免因小部件無效而引發錯誤。"""
        if self.log_messages and self.log_messages.winfo_exists():
            self.log_messages.configure(state='normal')  # 啟用編輯
            self.log_messages.insert(tk.END, message)
            self.log_messages.see(tk.END)  # 滾動到最後
            self.log_messages.configure(state='disabled')  # 禁用編輯
        else:
            print("Log Messages widget does not exist. Message:", message)

    def process_pr_no(self):
        """
        獲取並處理 PR No 及相關專案資料。
        """
        # 獲取 PR No
        return_block = self.get_attribute_block_with_name('pr_no')
        pr_no = self.get_block_text(return_block)
        if pr_no:
            self.pr_no = pr_no
            self.safe_log_insert(f"PR No: {pr_no}\n")

            # 使用 UtilOdoo 獲取專案資料
            project = self.odoo_util.get_project(pr_no)
            if project:
                self.project_id = project.get('id')
                self.project_name = project.get('name')
                self.job_working_plan_id = project.get('job_working_plan_id')
                self.job_working_plan_name = project.get('job_working_plan_name')
                self.safe_log_insert(f"專案ID: {self.project_id}\n")
                self.safe_log_insert(f"專案名稱: {self.project_name}\n")
                self.safe_log_insert(f"工種群組ID: {self.job_working_plan_id}\n")
                self.safe_log_insert(f"工種群組名稱: {self.job_working_plan_name}\n")

                # 設置塊屬性 project_name 和 job_working_plan_name
                _block = self.get_attribute_block()
                if _block:
                    self.set_attribute_value(_block, 'project_name', self.project_name)
                    self.set_attribute_value(_block, 'job_working_plan_name', self.job_working_plan_name)
            else:
                self.safe_log_insert("未找到對應的專案資料。\n")
        else:
            self.safe_log_insert("未找到名為 'pr_no' 的塊。\n")

    def get_active_layout(self):
        try:
            # Get active document
            doc = self.doc
            if not doc:
                raise Exception("No active document")

            # Get active layout
            layout = doc.ActiveLayout
            if not layout:
                raise Exception("No active layout")

            # Get layout properties
            layout_name = layout.Name
            
            # print(f"Active Layout Name: {layout_name}")
            self.safe_log_insert(f"Active Layout Name: {layout_name}\n")
            # print(f"Layout Block: {layout_block}")
            
            return layout

        except Exception as e:
            print(f"Error getting active layout: {str(e)}")
            return None

    def get_block_text(self, block):
        try:
            # Get block effective name
            block_name = block.EffectiveName
            
            # Get blocks collection from active document
            blocks = self.doc.Blocks
            
            # Get block definition
            block_def = blocks.Item(block_name)
            
            # Initialize pr_no
            pr_no = None
            
            # Iterate through items in block definition
            for item in block_def:
                # Check if item is a text object
                if item.ObjectName == "AcDbText":
                    pr_no = item.TextString
                    # self.log_messages.insert(tk.END, f"找到 pr_no: {pr_no}\n")
                    break
                    
            return pr_no
            
        except Exception as e:
            self.safe_log_insert(f"取得區塊文字時發生錯誤: {str(e)}\n")
            return None

    def get_attribute_block(self):
        """
        獲取特定的屬性塊。
        根據提供的塊名稱查找塊。
        
        :param block_name: 要查找的塊名稱
        :return: 塊對象或 None
        """
        try:
            layout = self.get_active_layout()
            if not layout:
                self.safe_log_insert("無法獲取活動佈局。\n")
                return None

            blocks = layout.Block
            for block in blocks:
                if block.ObjectName == "AcDbBlockReference":
                    self.safe_log_insert(f"找到屬性塊: {block.Name}\n")
                    attributes = block.GetAttributes()
                    for att in attributes:
                        if att.TagString == 'project_name' or att.TagString == 'job_working_plan_name':
                            return_block = block
            if not return_block:
                self.safe_log_insert("未找到屬性塊。\n")
                return None
            return return_block
        except Exception as e:
            self.safe_log_insert(f"獲取屬性塊時發生錯誤: {str(e)}\n")
            return None

    def get_attribute_block_with_name(self, block_name):
        """
        獲取特定的屬性塊。
        根據提供的塊名稱查找塊。
        
        :param block_name: 要查找的塊名稱
        :return: 塊對象或 None
        """
        try:
            layout = self.get_active_layout()
            if not layout:
                self.safe_log_insert("無法獲取活動佈局。\n")
                return None

            blocks = layout.Block
            for block in blocks:
                if block.ObjectName == "AcDbBlockReference" and block.Name == block_name:
                    self.safe_log_insert(f"找到屬性塊: {block.Name}\n")
                    return block
            self.safe_log_insert("未找到指定的屬性塊。\n")
            return None
        except Exception as e:
            self.safe_log_insert(f"獲取屬性塊時發生錯誤: {str(e)}\n")
            return None

    def set_attribute_values(self, block, attr_values):
        """
        設置指定塊的屬性值。

        :param block: AutoCAD 中的塊對象
        :param attr_values: 字典，包含屬性標籤及其對應的值，例如 {'TAG1': 'Value1', 'TAG2': 'Value2'}
        """
        try:
            attributes = block.GetAttributes()
            for att in attributes:
                tag = att.TagString
                if tag in attr_values:
                    old_value = att.TextString
                    att.TextString = attr_values[tag]
                    self.safe_log_insert(f"設置屬性 '{tag}' 從 '{old_value}' 為 '{attr_values[tag]}'\n")
            self.safe_log_insert("屬性設置完成。\n")
        except Exception as e:
            self.safe_log_insert(f"設置屬性值時發生錯誤: {str(e)}\n")

    def set_attribute_value(self, block, tag, val):
        """
        設置指定塊的單個屬性值。

        :param block: AutoCAD 中的塊對象
        :param tag: 屬性標籤 (不區分大小寫)
        :param val: 要設置的屬性值
        :return: 設置的值或 None 如果未找到匹配的屬性
        """
        try:
            tag_upper = tag.upper()
            print(f"tag_upper: {tag_upper}")
            attributes = block.GetAttributes()
            for att in attributes:
                att_tag_upper = att.TagString.upper()
                print(f"att_tag_upper: {att_tag_upper}")
                if att_tag_upper == tag_upper:
                    old_value = att.TextString
                    att.TextString = val
                    self.safe_log_insert(f"設置屬性 '{att.TagString}' 從 '{old_value}' 為 '{val}'\n")
                    return val
            self.safe_log_insert(f"未找到標籤為 '{tag}' 的屬性。\n")
            return None
        except Exception as e:
            self.safe_log_insert(f"設置屬性值時發生錯誤: {str(e)}\n")
            return None

    def clear_main_body(self, main_body):
        for widget in main_body.winfo_children():
            widget.destroy()
