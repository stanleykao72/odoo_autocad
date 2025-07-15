# -*- coding: utf-8 -*-
import tkinter as tk
from win32com import client
import pythoncom
import re
import time

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

            try:
                # 嘗試連接到已開啟的 AutoCAD 應用程序
                self.acad = client.GetActiveObject("AutoCAD.Application")
                self.log.safe_log_insert("已連接到現有的 AutoCAD 應用程序。\n")
            except client.pythoncom.com_error:
                # 如果未運行，則啟動 AutoCAD
                self.acad = client.Dispatch("AutoCAD.Application")
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
