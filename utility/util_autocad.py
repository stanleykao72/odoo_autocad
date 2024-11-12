# -*- coding: utf-8 -*-
import tkinter as tk
from win32com import client
import pythoncom

class UtilAutoCAD:
    def __init__(self):
        self.acad = None
        self.doc = None

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
                
                # 獲取 PR No
                pr_no, return_block = self.get_pr_no()
                if pr_no:
                    self.log_messages.insert(tk.END, f"PR No: {pr_no}\n")
                else:
                    self.log_messages.insert(tk.END, "未找到名為 'pr_no' 的塊。\n")
                
                main_body.after(1000, lambda: self.log_messages.insert(tk.END, "成功連接到 AutoCAD。\n"))
            else:
                main_body.after(1000, lambda: self.log_messages.insert(tk.END, "無法取得 AutoCAD 文件，請確認 AutoCAD 已開啟且有文件。\n"))
                
        except Exception as e:
            main_body.after(1000, lambda: self.log_messages.insert(tk.END, f"連接 AutoCAD 時發生錯誤: {str(e)}\n"))

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
            self.log_messages.insert(tk.END, f"Active Layout Name: {layout_name}\n")
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
            self.log_messages.insert(tk.END, f"取得區塊文字時發生錯誤: {str(e)}\n")
            return None

    def get_pr_no(self):
        pr_no = None
        return_block = None
        try:
            # 獲取當前佈局
            layout = self.get_active_layout()
            
            # print layout name in log
            # layout_name = layout.LayoutName
            # self.log_messages.insert(tk.END, f"Layout 名稱: {layout_name}\n")

            # 獲取佈局中的所有塊
            blocks = layout.Block
            
            for block in blocks:
                # self.log_messages.insert(tk.END, f"塊名稱: {block.Name}\n")
                if block.ObjectName == "AcDbBlockReference":
                    block_name = block.Name
                    self.log_messages.insert(tk.END, f"塊名稱: {block_name}\n")
                    if block_name == "pr_no":
                        self.log_messages.insert(tk.END, f"in if 塊名稱: {block_name}\n")
                        pr_no = self.get_block_text(block)
                        return_block = block
                        break
        except Exception as e:
            print(f"獲取 PR No 時發生錯誤: {str(e)}")
        
        return pr_no, return_block

    def clear_main_body(self, main_body):
        for widget in main_body.winfo_children():
            widget.destroy()
