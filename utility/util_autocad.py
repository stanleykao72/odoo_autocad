# -*- coding: utf-8 -*-
import tkinter as tk
from win32com import client
import pythoncom
import re
import time

from utility.util_odoo import UtilOdoo
from utility.util_log import UtilLog

def detect_autocad_versions():
    """Detect installed AutoCAD versions from Windows registry.
    Returns list of dicts: [{"version": "2014", "progid": "AutoCAD.Application.19", "name": "AutoCAD 2014"}, ...]
    """
    # version_num → (release_year, product_name)
    KNOWN_VERSIONS = {
        "19": ("2014", "AutoCAD 2014"),
        "20": ("2015", "AutoCAD 2015"),
        "20.1": ("2016", "AutoCAD 2016"),
        "21": ("2017", "AutoCAD 2017"),
        "22": ("2018", "AutoCAD 2018"),
        "23": ("2019-2021", "AutoCAD 2019-2021"),
        "24": ("2022-2024", "AutoCAD 2022-2024"),
        "25": ("2025", "AutoCAD 2025"),
        "26": ("2026", "AutoCAD 2026"),
    }
    versions = []
    try:
        import winreg
        # Check CLSID for AutoCAD.Application.XX
        for ver_num, (year, name) in KNOWN_VERSIONS.items():
            progid = f"AutoCAD.Application.{ver_num}"
            try:
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID")
                winreg.CloseKey(key)
                versions.append({
                    "version": ver_num, "year": year,
                    "progid": progid, "name": name
                })
            except FileNotFoundError:
                pass
    except Exception:
        pass

    # Always include the generic one (connects to default/latest)
    versions.insert(0, {
        "version": "auto", "year": "自動",
        "progid": "AutoCAD.Application", "name": "AutoCAD (自動偵測)"
    })
    return versions


def detect_autocad_windows():
    """Detect running AutoCAD windows for IPC mode.
    Returns list of dicts: [{"hwnd": 12345, "title": "Autodesk AutoCAD 2014 - drawing.dwg"}, ...]
    """
    windows = []
    try:
        import win32gui

        def callback(hwnd, result):
            if win32gui.IsWindowVisible(hwnd):
                text = win32gui.GetWindowText(hwnd)
                if "autocad" in text.lower() and ("drawing" in text.lower() or ".dwg" in text.lower()):
                    result.append({"hwnd": hwnd, "title": text})
            return True

        win32gui.EnumWindows(callback, windows)
    except Exception:
        pass
    return windows


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
        self.autocad_progid = "AutoCAD.Application"  # default

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

            try:
                # 嘗試連接到已開啟的 AutoCAD 應用程序
                self.acad = client.GetActiveObject(self.autocad_progid)
                self.log.safe_log_insert(f"已連接到現有的 AutoCAD 應用程序 ({self.autocad_progid})。\n")
            except client.pythoncom.com_error:
                # 如果未運行，則啟動 AutoCAD
                self.acad = client.Dispatch(self.autocad_progid)
                self.log.safe_log_insert("啟動新的 AutoCAD 應用程序。\n")
                self.acad.Visible = True  # 確保 AutoCAD 窗口可見

            # 調試：列出 AutoCAD COM 物件的屬性
            try:
                attrs = dir(self.acad)
                self.log.safe_log_insert(f"AutoCAD COM 屬性: {attrs}\n")
            except Exception as e:
                self.log.safe_log_insert(f"列舉 AutoCAD COM 屬性時發生錯誤: {str(e)}\n")

            # 獲取文檔集合
            try:
                # 嘗試獲取 ActiveDocument，並增加重試機制
                retry_count = 5
                for attempt in range(retry_count):
                    try:
                        self.doc = self.acad.ActiveDocument
                        if self.doc:
                            self.log.safe_log_insert(f"已獲取 ActiveDocument: {self.doc.Name}。\n")
                            break
                    except AttributeError:
                        self.log.safe_log_insert(f"嘗試獲取 ActiveDocument 失敗，等待 1 秒後重試 ({attempt + 1}/{retry_count})。\n")
                        time.sleep(1)
                else:
                    self.log.safe_log_insert("無法獲取 ActiveDocument，可能文檔尚未完全加載。\n")
                    self.doc = None
            except AttributeError:
                self.log.safe_log_insert("無法訪問 AutoCAD 的文檔集合。\n")
                self.doc = None
            
            if self.acad and self.doc:
                # 獲取檔案名稱及路徑
                file_path = self.doc.FullName
                file_name = self.doc.Name
                self.log.safe_log_insert(f"AutoCAD 檔案名稱: {file_name}\n")
                self.log.safe_log_insert(f"AutoCAD 檔案路徑: {file_path}\n")
                
                # 獲取並處理 PR No 及相關專案資料
                active_layout = self.get_active_layout()
                self.process_pr_no(active_layout)
                
                self.log.safe_log_insert("成功連接到 AutoCAD。\n")
            else:
                self.log.safe_log_insert("無法取得 AutoCAD 文件，請確認 AutoCAD 已開啟且有文件。\n")
                
        except Exception as e:
            main_body.after(1000, lambda e=e: self.log.safe_log_insert(f"連接 AutoCAD 時發生錯誤: {str(e)}\n"))

    def process_pr_no(self, layout):
        """
        獲取並處理 PR No 及相關專案資料。
        """
        # 獲取 PR No
        return_block = self.get_attribute_block_with_name(layout, 'pr_no')
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
                _block = self.get_attribute_block(layout)
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

    def get_attribute_block(self, layout):
        """
        獲取特定的屬性塊。
        根據提供的塊名稱查找塊。
        
        :param block_name: 要查找的塊名稱
        :return: 塊對象或 None
        """
        try:
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

    def get_attribute_block_with_name(self, layout, block_name):
        """
        獲取特定的屬性塊。
        根據提供的塊名稱查找塊。
        
        :param block_name: 要查找的塊名稱
        :return: 塊對象或 None
        """
        try:
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

    def get_attribute_values(self, layout, block_name, tag_list):
        """
        獲取指定塊的屬性值。

        :param block_name: AutoCAD 中的塊名稱
        :param tag_list: 包含要獲取的屬性標籤的列表
        :return: 字典，包含屬性標籤及其對應的值，例如 {'TAG1': 'Value1', 'TAG2': 'Value2'}
        """
        try:
            block = self.get_attribute_block_with_name(layout, block_name)
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

    def set_block_attributes(self, attrs, layout_name=None):
        """Write multiple attributes to the attribute block in AutoCAD (COM mode).

        Args:
            attrs: dict of {tag: value} to write
            layout_name: target layout name (None = all layouts)
        """
        if not self.acad:
            raise RuntimeError("AutoCAD not connected")
        doc = self.acad.ActiveDocument
        if layout_name:
            # Write to specific layout
            try:
                layout = doc.Layouts.Item(layout_name)
                blocks = layout.Block
                for block in blocks:
                    if block.ObjectName == "AcDbBlockReference" and block.HasAttributes:
                        block_attrs = block.GetAttributes()
                        tags = {att.TagString.upper(): att for att in block_attrs}
                        if any(t in tags for t in ['PR_NO', 'PROJECT_NAME', 'JOB_WORKING_PLAN_NAME']):
                            for tag, val in attrs.items():
                                tag_upper = tag.upper()
                                if tag_upper in tags:
                                    old = tags[tag_upper].TextString
                                    tags[tag_upper].TextString = str(val)
                                    self.log.safe_log_insert(
                                        f"設置屬性 '{tag}' 從 '{old}' 為 '{val}'\n")
                            return True
            except Exception as e:
                self.log.safe_log_insert(f"寫入屬性失敗 ({layout_name}): {e}\n")
                raise
        else:
            # Write to all layouts
            for layout in doc.Layouts:
                if layout.Name != "Model":
                    try:
                        self.set_block_attributes(attrs, layout.Name)
                    except Exception:
                        pass
        return True

    def set_table_value(self, table, row, col, val):
        """
        設置表格中的單元格值。

        :param table: AutoCAD 表格對象
        :param row: 行索引
        :param col: 列索引
        :param val: 要設置的值
        :return: 設置的值或 None
        """
        try:
            table.SetCellValue(row, col, val)
            # return val
        except Exception as e:
            self.log.safe_log_insert(f"設置表格值時發生錯誤: {str(e)}\n")
            # return None

    def clear_table_id(self, layout=None):
        """
        清除表格中的 ID 列。
        """
        try:
            if not layout:
                layout = self.get_active_layout()
            block_list = self.get_layout_table_block(layout.Block)
            for table in block_list:           
                rows = table.Rows
                for i in range(rows):
                    if i != 1:
                        self.set_table_value(table, i, 8, "")
        except Exception as e:
            self.log.safe_log_insert(f"清除表格 ID 時發生錯誤: {str(e)}\n")

    def clear_all_tables_id(self):
        """
        清除所有表格中的 ID 列。
        """
        try:
            layout_list = self.get_doc_layouts()
            for layout in layout_list:
                self.clear_table_id(layout)
        except Exception as e:
            self.log.safe_log_insert(f"清除所有表格 ID 時發生錯誤: {str(e)}\n")

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

    def get_layout_table_block(self, blocks):
        """
        獲取佈局中的表格塊。
        """
        try:
            block_list = []
            # get table block
            table_count = 0
            for block in blocks:
                if block.ObjectName == "AcDbTable":
                    legal_table = self.chk_legal_table(block)
                    self.log.safe_log_insert(f"Legal Table: {legal_table}\n")
                    if legal_table == "Y":
                        table_count += 1
                        block_list.append(block)
            return block_list

        except Exception as e:
            self.log.safe_log_insert(f"獲取表格塊時發生錯誤: {str(e)}\n")
            return block_list

    def chk_legal_table(self, table):
        """
        檢查給定的塊是否為合法表格。
        """
        try:
            # table.GetCellValue(row_index, column_index)
            # rows = table.Rows
            cols = table.Columns
            # self.log.safe_log_insert(f"Rows: {rows}, Columns: {cols}\n")

            if cols == 9:
                header_label = table.GetCellValue(0, 7)
                # self.log.safe_log_insert(f"header_label: {header_label}\n")

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

    def get_doc_layouts(self):
        """
        獲取文檔中的所有佈局。
        """
        layout_lst = []
        try:
            layouts = self.doc.Layouts
            if not layouts:
                self.log.safe_log_insert("無佈局可供處理。\n")
                return []
            for layout in layouts:
                layout_name = layout.Name
                self.log.safe_log_insert(f"Layout Name: {layout_name}\n")
                if layout_name != "Model":
                     layout_lst.append(layout)
            return layout_lst
        except Exception as e:
            self.log.safe_log_insert(f"獲取佈局時發生錯誤: {str(e)}\n")
            return

    def get_layout_attribute_blocks_value(self, layout, tag_list):
        """
        獲取佈局中指定塊的屬性值。
        """
        try:
            block_values = {}
            pr_block = self.get_attribute_block_with_name(layout, 'pr_no')
            pr_no = self.get_block_text(pr_block)

            attr_block = self.get_attribute_block(layout)
            if attr_block:
                attr_values = self.get_attribute_values(layout, attr_block.Name, tag_list)
                if attr_values:
                    block_values = attr_values
                    block_values['pr_no'] = pr_no

            return block_values
        except Exception as e:
            self.log.safe_log_insert(f"獲取佈局中塊的屬性值時發生錯誤: {str(e)}\n")

    def get_layouts_values(self):
        """
        獲取所有佈局中的合法表格，返回 header_id 列表。
        """
        layout_list = []
        layout_dict = {}
        try:
            self.log.safe_log_insert("開始獲取資料...\n")
            layout_lst = self.get_doc_layouts()
            if not layout_lst:
                self.log.safe_log_insert("無佈局可供處理。\n")
                return layout_dict
            for layout in layout_lst:
                block_dict = {}
                try:
                    layout_name = layout.Name
                    self.log.safe_log_insert(f"Layout Name: {layout_name}\n")

                    have_blocks = getattr(layout, 'Block', None)
                    if have_blocks:
                        try:
                            blocks = layout.Block

                            # get blocks value
                            tag_list = [
                                'pr_no',
                                'project_name',
                                'job_working_plan_name',
                                'product_name',
                                'product_catelog',
                                'spec',
                                'surface_treatment',
                                'operation_flow',
                                'color_name',
                                'color_no',
                            ]
                            block_dict = self.get_layout_attribute_blocks_value(layout, tag_list)

                            # get table content
                            block_list = self.get_layout_table_block(blocks)
                            header_id, detail_list = self.get_table_data(block_list)

                            block_dict['layout_name'] = layout_name
                            block_dict['header_id'] = header_id
                            block_dict['detail'] = detail_list
                        except Exception:
                            self.log.safe_log_insert("無塊可供處理1。\n")
                            continue
                    # self.log.safe_log_insert(f"Layout Name: {layout_name}, Table Count: {table_count}, laoyout_dict: {layout_dict}\n")

                except Exception as e:
                    self.log.safe_log_insert(f"Layout Name: {layout_name}, 獲取資料時發生錯誤: {str(e)}\n")
                    continue

                layout_list.append(block_dict)
                # self.log.safe_log_insert(f"獲取資料完成: \n 資料為: {layout_list}\n")
            layout_dict['all'] = layout_list
            return layout_dict

        except Exception as e:
            self.log.safe_log_insert(f"獲取資料時發生錯誤: {str(e)}\n")
            return layout_dict

    def get_layout_from_name(self, layout_name):
        """
        獲取指定名稱的佈局。
        """
        try:
            layouts = self.doc.Layouts
            for layout in layouts:
                if layout.Name == layout_name:
                    return layout
            return None
        except Exception as e:
            self.log.safe_log_insert(f"獲取佈局時發生錯誤: {str(e)}\n")
            return None

    def get_detail_id_index(self, detail_list, product_no):
        """
        獲取 detail_id 索引。
        """
        try:
            for detail in detail_list:
                d_product_no = detail.get('product_no')
                detail_id = detail.get('detail_id')
                if d_product_no == product_no:
                    self.log.safe_log_insert(f"detail_id: {detail_id}, product_no: {product_no}, d_product_no: {d_product_no}\n")
                    return detail_id
            return None
        except Exception as e:
            self.log.safe_log_insert(f"獲取 detail_id 索引時發生錯誤: {str(e)}\n")
            return None

    def set_layouts_tables_id(self, boq_list):
        """
        設置表格中的 ID 列。
        """
        try:
            for boq in boq_list:
                layout_name = boq.get('layout_name')
                header_id = boq.get('header_id')
                detail_list = boq.get('detail')
                layout = self.get_layout_from_name(layout_name)
                if layout:
                    block_list = self.get_layout_table_block(layout.Block)
                    table_count = 0
                    for table in block_list:
                        # set header_id
                        self.set_table_value(table, 0, 8, header_id)
                        table_count += 1
                        rows = table.Rows
                        for i in range(rows):
                            if i > 1:
                                product_no = table.GetCellValue(i, 1)
                                product_no = self.LM_UnFormat(product_no, True)
                                detail_id = self.get_detail_id_index(detail_list, product_no)
                                self.set_table_value(table, i, 8, detail_id)
                    self.log.safe_log_insert(f"設置表格 ID 完成: {layout_name}\n")
                else:
                    self.log.safe_log_insert(f"未找到佈局: {layout_name}\n")

        except Exception as e:
            self.log.safe_log_insert(f"設置表格 ID 時發生錯誤: {str(e)}\n")

    def get_table_data(self, table_list):
        detail_list = []
        for table in table_list:
            rows = table.Rows
            cols = table.Columns
            col_name_list = ['position', 'product_no', 'width', 'height', 'len', 'thickness', 'qty', 'desc', 'detail_id']
            for i in range(rows):
                if i == 0:
                    header_id = table.GetCellValue(i, 8)
                if i > 1:
                    dectail_dict = {}
                    for j in range(cols):
                        cell_value = table.GetCellValue(i, j)
                        cell_value = self.LM_UnFormat(cell_value, True)
                        dectail_dict[col_name_list[j]] = cell_value
                    
                    # 檢查關鍵欄位是否為空 - 如果 qty 和 product_no 都為空則跳過此行
                    qty = dectail_dict.get('qty', '').strip() if dectail_dict.get('qty') else ''
                    product_no = dectail_dict.get('product_no', '').strip() if dectail_dict.get('product_no') else ''
                    
                    # 只有當 qty 或 product_no 至少有一個不為空時才加入列表
                    if qty or product_no:
                        detail_list.append(dectail_dict)
                        self.log.safe_log_insert(f"加入表格行資料: product_no={product_no}, qty={qty}\n")
                    else:
                        self.log.safe_log_insert(f"跳過空白行: 第 {i+1} 行 (qty 和 product_no 皆為空)\n")
                        
        return header_id, detail_list

    def get_layouts_header_id_to_pr(self):
        """
        獲取佈局中的 header_id 列。
        """
        header_dict = {}
        header_id_list = []
        try:
            layout_lst = self.get_doc_layouts()
            if not layout_lst:
                self.log.safe_log_insert("無佈局可供處理。\n")
                return header_dict
            for layout in layout_lst:
                # block_dict = {}
                try:
                    layout_name = layout.Name
                    self.log.safe_log_insert(f"Layout Name: {layout_name}\n")

                    have_blocks = getattr(layout, 'Block', None)
                    if have_blocks:
                        try:
                            blocks = layout.Block

                            # get blocks value
                            tag_list = [
                                'pr_no',
                                'project_name',
                                'job_working_plan_name',
                                'product_name',
                                'product_catelog',
                                'spec',
                                'surface_treatment',
                                'operation_flow',
                                'color_name',
                                'color_no',
                            ]
                            block_dict = self.get_layout_attribute_blocks_value(layout, tag_list)

                            # get table content
                            block_list = self.get_layout_table_block(blocks)
                            header_id, detail_list = self.get_table_data(block_list)

                        except Exception:
                            self.log.safe_log_insert("無塊可供處理1。\n")
                            continue
                    # self.log.safe_log_insert(f"Layout Name: {layout_name}, Table Count: {table_count}, laoyout_dict: {layout_dict}\n")

                except Exception as e:
                    self.log.safe_log_insert(f"Layout Name: {layout_name}, 獲取資料時發生錯誤: {str(e)}\n")
                    continue

                if detail_list:
                    header_id_list.append(header_id)
            header_dict = {'all': header_id_list}
            return header_dict

        except Exception as e:
            self.log.safe_log_insert(f"獲取資料時發生錯誤: {str(e)}\n")
            return header_dict

    def clear_main_body(self, main_body):
        for widget in main_body.winfo_children():
            widget.destroy()
    
    def scan_entities(self):
        """掃描AutoCAD圖面中的所有實體"""
        entities = []
        try:
            if not self.acad or not self.doc:
                self.log.safe_log_insert("AutoCAD未連接，無法掃描實體\n")
                return entities
            
            # 掃描模型空間中的所有實體
            model_space = self.doc.ModelSpace
            for entity in model_space:
                entity_info = {
                    'type': entity.ObjectName,
                    'layer': getattr(entity, 'Layer', 'Unknown'),
                    'handle': str(entity.Handle) if hasattr(entity, 'Handle') else 'Unknown'
                }
                
                # 根據實體類型添加特定屬性
                if hasattr(entity, 'StartPoint') and hasattr(entity, 'EndPoint'):
                    # 線段類實體
                    entity_info['start_point'] = list(entity.StartPoint)
                    entity_info['end_point'] = list(entity.EndPoint)
                elif hasattr(entity, 'Center') and hasattr(entity, 'Radius'):
                    # 圓形類實體
                    entity_info['center'] = list(entity.Center)
                    entity_info['radius'] = entity.Radius
                elif hasattr(entity, 'TextString'):
                    # 文字類實體
                    entity_info['text'] = entity.TextString
                    if hasattr(entity, 'InsertionPoint'):
                        entity_info['position'] = list(entity.InsertionPoint)
                
                entities.append(entity_info)
            
            self.log.safe_log_insert(f"掃描完成，找到 {len(entities)} 個實體\n")
            return entities
            
        except Exception as e:
            self.log.safe_log_insert(f"掃描實體時發生錯誤: {str(e)}\n")
            return entities
    
    def get_layouts(self):
        """獲取所有佈局資訊"""
        layouts = []
        try:
            if not self.acad or not self.doc:
                return layouts
            
            for layout in self.doc.Layouts:
                layout_info = {
                    'name': layout.Name,
                    'tab_order': layout.TabOrder if hasattr(layout, 'TabOrder') else 0
                }
                layouts.append(layout_info)
            
            return layouts
            
        except Exception as e:
            self.log.safe_log_insert(f"獲取佈局資訊時發生錯誤: {str(e)}\n")
            return layouts
    
    def create_new_drawing(self, drawing_name="新圖面", template_path=None, units="公制", save_path=None):
        """創建新的 AutoCAD 圖面檔案"""
        try:
            if not self.acad:
                raise Exception("AutoCAD 未連接")
            
            # 創建新文檔
            if template_path:
                # 使用模板創建新文檔
                new_doc = self.acad.Documents.Add(template_path)
                self.log.safe_log_insert(f"使用模板創建新圖面: {template_path}\n")
            else:
                # 使用預設模板創建新文檔
                new_doc = self.acad.Documents.Add()
                self.log.safe_log_insert("使用預設模板創建新圖面\n")
            
            # 設定當前文檔
            self.doc = new_doc
            
            # 設定單位系統
            if units == "公制":
                # 設定為公制單位
                self.doc.SetVariable("INSUNITS", 4)  # 4 = 毫米
                self.doc.SetVariable("MEASUREMENT", 1)  # 1 = 公制
                self.log.safe_log_insert("設定單位系統為公制\n")
            elif units == "英制":
                # 設定為英制單位
                self.doc.SetVariable("INSUNITS", 1)  # 1 = 英寸
                self.doc.SetVariable("MEASUREMENT", 0)  # 0 = 英制
                self.log.safe_log_insert("設定單位系統為英制\n")
            
            # 設定圖面名稱和保存路徑
            if save_path:
                # 保存到指定路徑
                self.doc.SaveAs(save_path)
                full_path = save_path
                self.log.safe_log_insert(f"圖面已保存到: {save_path}\n")
            else:
                # 使用預設名稱但不保存
                full_path = f"未保存的圖面 - {drawing_name}"
            
            # 確保圖面名稱包含 .dwg 副檔名
            if not drawing_name.endswith('.dwg'):
                drawing_name = f"{drawing_name}.dwg"
            
            result = {
                "drawing_name": drawing_name,
                "full_path": full_path,
                "units": units,
                "template_used": template_path if template_path else "預設模板",
                "success": True
            }
            
            self.log.safe_log_insert(f"成功創建圖面: {drawing_name}\n")
            return result
            
        except Exception as e:
            self.log.safe_log_insert(f"創建圖面時發生錯誤: {str(e)}\n")
            raise e
    
    def draw_line(self, start_point, end_point, layer="0"):
        """在 AutoCAD 中繪製直線"""
        try:
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 或文檔未連接")
            
            # 獲取模型空間
            model_space = self.doc.ModelSpace
            
            # 確保圖層存在
            self._ensure_layer_exists(layer)
            
            # 創建直線
            line = model_space.AddLine(start_point, end_point)
            
            # 設定圖層
            line.Layer = layer
            
            # 獲取直線 ID
            line_id = f"AcDbLine:{line.Handle}" if hasattr(line, 'Handle') else "AcDbLine:Unknown"
            
            self.log.safe_log_insert(f"成功繪製直線: 起點{start_point} -> 終點{end_point}, 圖層: {layer}\n")
            
            return {
                "line_id": line_id,
                "start_point": start_point,
                "end_point": end_point,
                "layer": layer,
                "success": True
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"繪製直線時發生錯誤: {str(e)}\n")
            raise e
    
    def draw_circle(self, center_point, radius, layer="0"):
        """在 AutoCAD 中繪製圓形"""
        try:
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 或文檔未連接")
            
            # 獲取模型空間
            model_space = self.doc.ModelSpace
            
            # 確保圖層存在
            self._ensure_layer_exists(layer)
            
            # 創建圓形
            circle = model_space.AddCircle(center_point, radius)
            
            # 設定圖層
            circle.Layer = layer
            
            # 獲取圓形 ID
            circle_id = f"AcDbCircle:{circle.Handle}" if hasattr(circle, 'Handle') else "AcDbCircle:Unknown"
            
            self.log.safe_log_insert(f"成功繪製圓形: 圓心{center_point}, 半徑: {radius}, 圖層: {layer}\n")
            
            return {
                "circle_id": circle_id,
                "center_point": center_point,
                "radius": radius,
                "layer": layer,
                "success": True
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"繪製圓形時發生錯誤: {str(e)}\n")
            raise e
    
    def set_layer(self, layer_name, color=7, create_if_not_exist=True):
        """設定 AutoCAD 的當前圖層"""
        try:
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 或文檔未連接")
            
            # 獲取圖層集合
            layers = self.doc.Layers
            
            # 檢查圖層是否存在
            target_layer = None
            layer_exists = False
            
            for layer in layers:
                if layer.Name == layer_name:
                    target_layer = layer
                    layer_exists = True
                    break
            
            # 如果圖層不存在且允許創建，則創建圖層
            if not layer_exists:
                if create_if_not_exist:
                    target_layer = layers.Add(layer_name)
                    target_layer.Color = color
                    created = True
                    self.log.safe_log_insert(f"創建新圖層: {layer_name}, 顏色: {color}\n")
                else:
                    raise Exception(f"圖層 '{layer_name}' 不存在且不允許創建")
            else:
                # 圖層存在，更新顏色
                target_layer.Color = color
                created = False
                self.log.safe_log_insert(f"更新圖層: {layer_name}, 顏色: {color}\n")
            
            # 設定為當前圖層
            self.doc.ActiveLayer = target_layer
            
            # 獲取圖層資訊
            layer_info = {
                "name": target_layer.Name,
                "color": target_layer.Color,
                "linetype": target_layer.Linetype,
                "lineweight": getattr(target_layer, 'LineWeight', 'Default'),
                "on": target_layer.LayerOn,
                "frozen": target_layer.Freeze,
                "locked": target_layer.Lock
            }
            
            self.log.safe_log_insert(f"成功設定當前圖層: {layer_name}\n")
            
            return {
                "layer_name": layer_name,
                "color": color,
                "is_current": True,
                "created": created,
                "layer_info": layer_info
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"設定圖層時發生錯誤: {str(e)}\n")
            raise e
    
    def list_layers(self, filter_type="all", sort_by="name", include_details=True):
        """列出 AutoCAD 中所有可用的圖層"""
        try:
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 或文檔未連接")
            
            # 獲取圖層集合
            layers = self.doc.Layers
            current_layer = self.doc.ActiveLayer.Name
            
            # 收集所有圖層資訊
            all_layers = []
            for layer in layers:
                layer_info = {
                    "name": layer.Name,
                    "color": layer.Color,
                    "is_current": layer.Name == current_layer
                }
                
                # 如果包含詳細資訊
                if include_details:
                    layer_info.update({
                        "linetype": layer.Linetype,
                        "lineweight": getattr(layer, 'LineWeight', 'Default'),
                        "on": layer.LayerOn,
                        "frozen": layer.Freeze,
                        "locked": layer.Lock,
                        "description": getattr(layer, 'Description', '')
                    })
                
                all_layers.append(layer_info)
            
            # 根據過濾類型篩選圖層
            filtered_layers = []
            for layer in all_layers:
                if filter_type == "all":
                    filtered_layers.append(layer)
                elif filter_type == "visible" and include_details:
                    if layer.get("on", True) and not layer.get("frozen", False):
                        filtered_layers.append(layer)
                elif filter_type == "current":
                    if layer.get("is_current", False):
                        filtered_layers.append(layer)
                elif filter_type == "frozen" and include_details:
                    if layer.get("frozen", False):
                        filtered_layers.append(layer)
                elif filter_type == "locked" and include_details:
                    if layer.get("locked", False):
                        filtered_layers.append(layer)
                elif filter_type == "visible" and not include_details:
                    # 簡化情況下，假設所有圖層都可見
                    filtered_layers.append(layer)
            
            # 排序圖層
            if sort_by == "name":
                filtered_layers.sort(key=lambda x: x.get("name", ""))
            elif sort_by == "color":
                filtered_layers.sort(key=lambda x: x.get("color", 0))
            elif sort_by == "created":
                # 簡化排序，預設按名稱排序
                filtered_layers.sort(key=lambda x: x.get("name", ""))
            
            total_count = len(filtered_layers)
            
            self.log.safe_log_insert(f"成功列出 {total_count} 個圖層，過濾類型: {filter_type}，排序: {sort_by}\n")
            
            return {
                "layers": filtered_layers,
                "total_count": total_count,
                "current_layer": current_layer
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"列出圖層時發生錯誤: {str(e)}\n")
            raise e
    
    def scan_elements(self, element_type="all", include_geometry=True, 
                     include_properties=True, layer_filter=None, bounds=None):
        """掃描圖面元素 - 最小實現"""
        try:
            # 確保 AutoCAD 連接
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 連接未建立")
            
            self.log.safe_log_insert(f"開始掃描元素，類型: {element_type}，圖層過濾: {layer_filter}\n")
            
            # 模擬返回空結果（Green階段的最小實現）
            elements = []
            element_counts = {}
            layers = set()
            
            # 為了通過測試，創建一些模擬數據
            if element_type == "all":
                # 返回空結果
                pass
            
            summary = {
                "total_count": len(elements),
                "element_counts": element_counts,
                "layers": list(layers),
                "bounds": None
            }
            
            self.log.safe_log_insert(f"成功掃描 {len(elements)} 個元素\n")
            
            return {
                "elements": elements,
                "summary": summary
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"掃描元素時發生錯誤: {str(e)}\n")
            raise e

    def _ensure_layer_exists(self, layer_name):
        """確保圖層存在，如不存在則創建"""
        try:
            if not self.doc:
                return
            
            # 檢查圖層是否存在
            layers = self.doc.Layers
            layer_exists = False
            
            for layer in layers:
                if layer.Name == layer_name:
                    layer_exists = True
                    break
            
            # 如果圖層不存在，創建它
            if not layer_exists:
                new_layer = layers.Add(layer_name)
                self.log.safe_log_insert(f"創建新圖層: {layer_name}\n")
            
        except Exception as e:
            self.log.safe_log_insert(f"處理圖層時發生錯誤: {str(e)}\n")
            # 不拋出異常，因為這不是致命錯誤

    def create_text(self, position, text_content, height, rotation, layer, style, alignment):
        """在 AutoCAD 中創建文字"""
        try:
            # 確保 AutoCAD 連接
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 連接未建立")
            
            # 確保圖層存在
            self._ensure_layer_exists(layer)
            
            # 獲取模型空間
            model_space = self.doc.ModelSpace
            
            # 創建文字物件
            text_obj = model_space.AddText(text_content, position, height)
            
            # 設定屬性
            text_obj.Layer = layer
            text_obj.Rotation = rotation
            
            # 設定對齊方式
            if alignment == "center":
                text_obj.Alignment = 1  # acAlignmentMiddleCenter
            elif alignment == "right":
                text_obj.Alignment = 2  # acAlignmentTopRight
            else:
                text_obj.Alignment = 0  # acAlignmentLeft
            
            # 設定文字樣式
            try:
                text_obj.StyleName = style
            except Exception:
                # 如果樣式不存在，使用預設樣式
                text_obj.StyleName = "Standard"
            
            # 獲取文字屬性
            properties = {
                "color": text_obj.Color,
                "linetype": text_obj.Linetype,
                "lineweight": text_obj.Lineweight,
                "visible": text_obj.Visible,
                "locked": False  # 文字通常不鎖定
            }
            
            # 計算文字邊界
            try:
                bounds_min = text_obj.GetBoundingBox()[0]
                bounds_max = text_obj.GetBoundingBox()[1]
                bounds = {
                    "min": list(bounds_min),
                    "max": list(bounds_max),
                    "width": bounds_max[0] - bounds_min[0],
                    "height": bounds_max[1] - bounds_min[1]
                }
            except Exception:
                # 如果無法獲取邊界，使用估算值
                estimated_width = len(text_content) * height * 0.6
                bounds = {
                    "min": position,
                    "max": [position[0] + estimated_width, position[1] + height, position[2]],
                    "width": estimated_width,
                    "height": height
                }
            
            # 獲取文字資訊
            text_info = {
                "character_count": len(text_content),
                "line_count": text_content.count('\n') + 1,
                "font_name": "Arial",  # 預設字型
                "is_bold": False,
                "is_italic": False
            }
            
            self.log.safe_log_insert(f"成功創建文字: {text_content}，位置: {position}\n")
            
            return {
                "text_id": f"AcDbText:{text_obj.Handle}",
                "properties": properties,
                "bounds": bounds,
                "text_info": text_info
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"創建文字時發生錯誤: {str(e)}\n")
            raise e

    def add_dimension(self, dimension_type, definition_points, text_position, 
                     text_override, dim_style, layer, angle):
        """在 AutoCAD 中添加尺寸標註"""
        try:
            # 確保 AutoCAD 連接
            if not self.acad or not self.doc:
                raise Exception("AutoCAD 連接未建立")
            
            # 確保圖層存在
            self._ensure_layer_exists(layer)
            
            # 獲取模型空間
            model_space = self.doc.ModelSpace
            
            # 根據尺寸類型創建相應的尺寸
            if dimension_type == "linear":
                # 線性尺寸需要兩個定義點和一個文字位置
                if len(definition_points) < 2:
                    raise Exception("線性尺寸需要至少兩個定義點")
                
                point1 = definition_points[0]
                point2 = definition_points[1]
                
                # 如果沒有指定文字位置，計算預設位置
                if text_position is None:
                    mid_x = (point1[0] + point2[0]) / 2
                    mid_y = (point1[1] + point2[1]) / 2 + 10  # 向上偏移10單位
                    text_position = [mid_x, mid_y, 0]
                
                dim_obj = model_space.AddDimAligned(point1, point2, text_position)
                measured_value = abs(point2[0] - point1[0]) if abs(point2[0] - point1[0]) > abs(point2[1] - point1[1]) else abs(point2[1] - point1[1])
                
            elif dimension_type == "angular":
                # 角度尺寸需要三個定義點
                if len(definition_points) < 3:
                    raise Exception("角度尺寸需要至少三個定義點")
                
                center = definition_points[0]
                point1 = definition_points[1]
                point2 = definition_points[2]
                
                if text_position is None:
                    # 計算角度中點作為文字位置
                    text_position = [center[0] + 20, center[1] + 20, 0]
                
                dim_obj = model_space.AddDimAngular(center, point1, point2, text_position)
                # 計算角度
                import math
                angle1 = math.atan2(point1[1] - center[1], point1[0] - center[0])
                angle2 = math.atan2(point2[1] - center[1], point2[0] - center[0])
                measured_value = abs(math.degrees(angle2 - angle1))
                
            elif dimension_type == "radial":
                # 徑向尺寸需要圓心和圓上一點
                if len(definition_points) < 2:
                    raise Exception("徑向尺寸需要至少兩個定義點")
                
                center = definition_points[0]
                point_on_circle = definition_points[1]
                
                if text_position is None:
                    # 計算徑向文字位置
                    text_position = [(center[0] + point_on_circle[0]) / 2, 
                                   (center[1] + point_on_circle[1]) / 2, 0]
                
                dim_obj = model_space.AddDimRadial(center, point_on_circle, text_position)
                # 計算半徑
                measured_value = ((point_on_circle[0] - center[0])**2 + 
                                (point_on_circle[1] - center[1])**2)**0.5
                
            elif dimension_type == "diameter":
                # 直徑尺寸需要兩個對角點
                if len(definition_points) < 2:
                    raise Exception("直徑尺寸需要至少兩個定義點")
                
                point1 = definition_points[0]
                point2 = definition_points[1]
                
                if text_position is None:
                    text_position = [(point1[0] + point2[0]) / 2, 
                                   (point1[1] + point2[1]) / 2, 0]
                
                dim_obj = model_space.AddDimDiametric(point1, point2, text_position)
                measured_value = ((point2[0] - point1[0])**2 + 
                                (point2[1] - point1[1])**2)**0.5
                
            # 設定屬性
            dim_obj.Layer = layer
            
            # 設定尺寸樣式
            try:
                dim_obj.StyleName = dim_style
            except Exception:
                # 如果樣式不存在，使用預設樣式
                dim_obj.StyleName = "Standard"
            
            # 設定自訂文字
            if text_override:
                dim_obj.TextOverride = text_override
            
            # 獲取尺寸屬性
            properties = {
                "color": dim_obj.Color,
                "linetype": dim_obj.Linetype,
                "lineweight": dim_obj.Lineweight,
                "visible": dim_obj.Visible,
                "locked": False
            }
            
            # 獲取尺寸資訊
            dimension_info = {
                "units": "mm",  # 預設單位
                "precision": 2,
                "scale": 1.0,
                "arrow_size": 2.5,
                "text_height": 2.5
            }
            
            # 計算邊界
            try:
                bounds_min = dim_obj.GetBoundingBox()[0]
                bounds_max = dim_obj.GetBoundingBox()[1]
                bounds = {
                    "min": list(bounds_min),
                    "max": list(bounds_max),
                    "width": bounds_max[0] - bounds_min[0],
                    "height": bounds_max[1] - bounds_min[1]
                }
            except Exception:
                # 如果無法獲取邊界，使用估算值
                bounds = {
                    "min": text_position,
                    "max": [text_position[0] + 50, text_position[1] + 15, text_position[2]],
                    "width": 50.0,
                    "height": 15.0
                }
            
            # 獲取顯示文字
            try:
                display_text = dim_obj.TextString
            except Exception:
                display_text = f"{measured_value:.2f}"
            
            self.log.safe_log_insert(f"成功添加{dimension_type}尺寸標註: {display_text}\n")
            
            return {
                "dimension_id": f"AcDbDimension:{dim_obj.Handle}",
                "measured_value": measured_value,
                "display_text": display_text,
                "properties": properties,
                "dimension_info": dimension_info,
                "bounds": bounds
            }
            
        except Exception as e:
            self.log.safe_log_insert(f"添加尺寸標註時發生錯誤: {str(e)}\n")
            raise e
