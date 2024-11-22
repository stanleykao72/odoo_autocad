# -*- coding: utf-8 -*-
import tkinter as tk
from win32com import client
import pythoncom
import re

from utility.util_odoo import UtilOdoo
from utility.util_log import UtilLog

class UtilAutoCAD:
    def __init__(self, odoo_util, log_util):
        self.acad = None
        self.doc = None
        self.pr_no = None
        self.project_id = None
        self.project_name = None
        self.job_working_plan_id = None
        self.job_working_plan_name = None
        self.log = log_util
        self.odoo_util = odoo_util

    def connected_autocad(self):
        if self.acad:
            return True
        else:
            return False

    def connect_autocad(self, main_body):
        self.clear_main_body(main_body)
        self.log.safe_log_insert("連接到 AutoCAD...\n")
        
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
                self.log.safe_log_insert(f"AutoCAD 檔案名稱: {file_name}\n")
                self.log.safe_log_insert(f"AutoCAD 檔案路徑: {file_path}\n")
                
                # 獲取並處理 PR No 及相關專案資料
                self.process_pr_no()
                
                self.log.safe_log_insert("成功連接到 AutoCAD。\n")
            else:
                self.log.safe_log_insert("無法取得 AutoCAD 文件，請確認 AutoCAD 已開啟且有文件。\n")
                
        except Exception as e:
            main_body.after(1000, lambda e=e: self.log.safe_log_insert(f"連接 AutoCAD 時發生錯誤: {str(e)}\n"))

    def process_pr_no(self):
        """
        獲取並處理 PR No 及相關專案資料。
        """
        # 獲取 PR No
        return_block = self.get_attribute_block_with_name('pr_no')
        pr_no = self.get_block_text(return_block)
        if pr_no:
            self.pr_no = pr_no
            self.log.safe_log_insert(f"PR No: {pr_no}\n")

            # 使用 UtilOdoo 獲取專案資料
            project = self.odoo_util.get_project(pr_no)
            if project:
                self.project_id = project.get('id')
                self.project_name = project.get('name')
                self.job_working_plan_id = project.get('job_working_plan_id')
                self.job_working_plan_name = project.get('job_working_plan_name')
                self.log.safe_log_insert(f"專案ID: {self.project_id}\n")
                self.log.safe_log_insert(f"專案名稱: {self.project_name}\n")
                self.log.safe_log_insert(f"工種群組ID: {self.job_working_plan_id}\n")
                self.log.safe_log_insert(f"工種群組名稱: {self.job_working_plan_name}\n")

                # 設置塊屬性 project_name 和 job_working_plan_name
                _block = self.get_attribute_block()
                if _block:
                    self.set_attribute_value(_block, 'project_name', self.project_name)
                    self.set_attribute_value(_block, 'job_working_plan_name', self.job_working_plan_name)
            else:
                self.log.safe_log_insert("未找到對應的專案資料。\n")
        else:
            self.log.safe_log_insert("未找到名為 'pr_no' 的塊。\n")

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
            self.log.safe_log_insert(f"Active Layout Name: {layout_name}\n")
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
            block_text = None
            
            # Iterate through items in block definition
            for item in block_def:
                # Check if item is a text object
                if item.ObjectName == "AcDbText":
                    block_text = item.TextString
                    break
            return block_text
            
        except Exception as e:
            self.log.safe_log_insert(f"取得區塊文字時發生錯誤: {str(e)}\n")
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
                self.log.safe_log_insert("無法獲取活動佈局。\n")
                return None

            blocks = layout.Block
            for block in blocks:
                if block.ObjectName == "AcDbBlockReference":
                    self.log.safe_log_insert(f"找到屬性塊: {block.Name}\n")
                    attributes = block.GetAttributes()
                    for att in attributes:
                        if att.TagString == 'project_name' or att.TagString == 'job_working_plan_name':
                            return_block = block
            if not return_block:
                self.log.safe_log_insert("未找到屬性塊。\n")
                return None
            return return_block
        except Exception as e:
            self.log.safe_log_insert(f"獲取屬性塊時發生錯誤: {str(e)}\n")
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
                self.log.safe_log_insert("無法獲取活動佈局。\n")
                return None

            blocks = layout.Block
            for block in blocks:
                if block.ObjectName == "AcDbBlockReference" and block.Name == block_name:
                    self.log.safe_log_insert(f"找到屬性塊: {block.Name}\n")
                    return block
            self.log.safe_log_insert("未找到指定的屬性塊。\n")
            return None
        except Exception as e:
            self.log.safe_log_insert(f"獲取屬性塊時發生錯誤: {str(e)}\n")
            return None

    def get_attribute_values(self, block_name, tag_list):
        """
        獲取指定塊的屬性值。

        :param block_name: AutoCAD 中的塊名稱
        :param tag_list: 包含要獲取的屬性標籤的列表
        :return: 字典，包含屬性標籤及其對應的值，例如 {'TAG1': 'Value1', 'TAG2': 'Value2'}
        """
        try:
            block = self.get_attribute_block_with_name(block_name)
            if block:
                attributes = block.GetAttributes()
                attr_values = {}
                for att in attributes:
                    if att.TagString in tag_list:
                        attr_values[att.TagString] = att.TextString
                return attr_values
            else:
                return None
        except Exception as e:
            self.log.safe_log_insert(f"獲取屬性值時發生錯誤: {str(e)}\n")
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
                    self.log.safe_log_insert(f"設置屬性 '{tag}' 從 '{old_value}' 為 '{attr_values[tag]}'\n")
            self.log.safe_log_insert("屬性設置完成。\n")
        except Exception as e:
            self.log.safe_log_insert(f"設置屬性值時發生錯誤: {str(e)}\n")

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
                    self.log.safe_log_insert(f"設置屬性 '{att.TagString}' 從 '{old_value}' 為 '{val}'\n")
                    return val
            self.log.safe_log_insert(f"未找到標籤為 '{tag}' 的屬性。\n")
            return None
        except Exception as e:
            self.log.safe_log_insert(f"設置屬性值時發生錯誤: {str(e)}\n")
            return None

    def LM_UnFormat(self, s, mtx):
        """
        根據提供的模式對字符串進行未格式化處理。

        :param s: 要處理的字符串
        :param mtx: 布爾值，決定是否進行額外的替換
        :return: 處理後的字符串或 None
        """
        try:
            # 先判斷 s 是否為空, 如果是則返回 None
            if not s:
                return False
            
            # 判斷 s 是否為字符串, 如果不是再判斷是否為 int 或 float, 如果都不是則返回 None
            if not isinstance(s, str):
                if isinstance(s, (int, float)):
                    return str(s)
                else:
                    return False

            # 定義替換規則，按照順序進行替換
            replacements = [
                (r"\\\\", "\032"),  # 替換 "\\\\" 為 "\032"
                (r"\\P|\n|\t", " "),  # 替換 "\P", 換行, 制表符 為 空格
                (r"\\(\\[ACcFfHLlOopQTW])|\\[ACcFfHLlOopQTW][^\\;]*;|\\[ACcFfHLlOopQTW]", ""),
                (r"([^\\])\\S([^;]*)[/#\^]([^;]*);", ""),
                (r"\\(\S)|[\\]({)|}", ""),
                (r"[\\]({)|{", ""),
            ]

            # 依序應用所有替換規則
            for pattern, repl in replacements:
                s = re.sub(pattern, repl, s)

            if mtx:
                # 額外的替換操作
                s = re.sub(r"(\\[ACcFfHLlOoPpQSTW])|{|}", "\\\\", s)
                s = s.replace("\\", "\032")
            else:
                s = s.replace("\\", "\032")
            
            return s

        except Exception as e:
            # 可根據需求處理異常，例如記錄日誌
            print(f"處理時發生錯誤: {e} -- {s}")
            return False

    def chk_legal_table(self, block):
        """
        檢查給定的塊是否為合法表格。
        """
        table = block
        try:
            # table.GetCellValue(row_index, column_index)
            # rows = table.Rows
            cols = table.Columns
            # self.log.safe_log_insert(f"Rows: {rows}, Columns: {cols}\n")

            if cols == 9:
                header_label = table.GetCellValue(0, 7)
                self.log.safe_log_insert(f"header_label: {header_label}\n")

                if header_label == "HEADER_ID":
                    return_str = "Y"
                else:
                    return_str = "N"
            else:
                return_str = "N"

            return return_str

        except Exception as e:
            self.log.safe_log_insert(f"檢查合法表格時發生錯誤: {str(e)}\n")
            return "N"

    def get_tables_from_layouts(self):
        """
        獲取所有佈局中的合法表格，返回 header_id 列表。
        """
        table_lst = []
        try:
            layouts = self.doc.Layouts
            if not layouts:
                self.log.safe_log_insert("無佈局可供處理。\n")
                return table_lst

            self.log.safe_log_insert("開始獲取表格塊...\n")
            for layout in layouts:
                try:
                    layout_name = layout.Name
                    if layout_name != "Model":
                        self.log.safe_log_insert(f"Layout Name: {layout_name}\n")

                        have_blocks = getattr(layout, 'Block', None)
                        if have_blocks:
                            try:
                                blocks = layout.Block
                            except Exception:
                                self.log.safe_log_insert("無塊可供處理。\n")
                                continue

                            if not blocks:
                                self.log.safe_log_insert("無塊可供處理。\n")
                                continue

                            table_count = 0
                            for block in blocks:
                                if block.ObjectName == "AcDbTable":
                                    legal_table = self.chk_legal_table(block)
                                    self.log.safe_log_insert(f"Legal Table: {legal_table}\n")
                                    if legal_table == "Y":
                                        table_count += 1
                                        table_lst.append({'layout': layout, 'block': block})
                except Exception as e:
                    self.log.safe_log_insert(f"Layout Name: {layout_name}, 獲取表格時發生錯誤: {str(e)}\n")
                    continue
            return table_lst

        except Exception as e:
            self.log.safe_log_insert(f"獲取表格塊時發生錯誤: {str(e)}\n")
            return table_lst

    def get_table_data(self, layout, table):
        rows = table.Rows
        cols = table.Columns
        col_name_list = ['position', 'product_no', 'width', 'height', 'lenght', 'thickness', 'qty', 'desc', 'detail_id']
        detail_list = []
        for i in range(rows):
            if i == 0:
                header_id = table.GetCellValue(i, 8)
            if i > 1:
                dectail_dict = {}
                for j in range(cols):
                    cell_value = table.GetCellValue(i, j)
                    cell_value = self.LM_UnFormat(cell_value, True)
                    # self.log.safe_log_insert(f"col_name_list[j]: {col_name_list[j]}\n")
                    # self.log.safe_log_insert(f"cell_value: {cell_value}\n")
                    dectail_dict[col_name_list[j]] = cell_value
                detail_list.append(dectail_dict)
        return_dict = {
            'layout': layout.Name,
            'header_id': header_id,
            'details': detail_list,
        }
        return return_dict

    def clear_main_body(self, main_body):
        for widget in main_body.winfo_children():
            widget.destroy()
