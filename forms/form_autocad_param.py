# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
from utility.util_autocad import UtilAutoCAD
# from utility.util_log import UtilLog
from utility.popup_selector import PopupSelector  # 新增導入


class FormAutoCADParam:
    def __init__(self, main_content, odoo_util, autocad_util, log_util, root):
        self.main_content = main_content
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util  # Store the UtilAutoCAD instance
        self.project_id = self.autocad_util.project_id
        self.job_working_plan_name = self.autocad_util.job_working_plan_name
        self.material_selector = None
        self.spec_selector = None
        self.category_selector = None

        self.util_log = log_util
        self.root = root

    def get_parameters_from_odoo(self):
        self.clear_main_content()
        form_layout = tk.Frame(self.main_content)
        form_layout.pack(fill="both", expand=True)

        # 建立一個帶有框線和標題的框架
        material_frame = tk.LabelFrame(form_layout, text="材料", padx=10, pady=10)
        material_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # 材料輸入區
        row1 = tk.Frame(material_frame)
        row1.pack(fill="x", pady=5)

        tk.Label(row1, text="材料:", width=15, anchor='w').pack(side="left")
        self.material_entry = tk.Entry(row1, justify='left')
        self.material_entry.pack(side="left", fill="x", expand=True, padx=5)

        self.material_uom_entry = tk.Entry(row1, width=8, state='readonly', justify='left')
        self.material_uom_entry.pack(side="left", padx=5)

        # 建立材料選擇器
        self.material_selector = PopupSelector(
            self.main_content,
            self.odoo_util.get_product,
            self.material_entry,
            ["name", "uom"],
            "選擇材料",
            additional_target_entries={"uom": self.material_uom_entry}
        )
        self.material_entry.bind("<FocusIn>",
                                 lambda e: self.material_selector.show())

        # 材質輸入區
        row2 = tk.Frame(material_frame)
        row2.pack(fill="x", pady=5)

        tk.Label(row2, text="材質:", width=15, anchor='w').pack(side="left")
        self.spec_entry = tk.Entry(row2, justify='left')
        self.spec_entry.pack(side="left", fill="x", expand=True, padx=5)

        # 建立材質選擇器
        self.spec_selector = PopupSelector(
            self.main_content,
            lambda: self.odoo_util.get_setup('spec'),
            self.spec_entry,
            ["value"],
            "選擇材質"
        )
        self.spec_entry.bind("<FocusIn>",
                             lambda e: self.spec_selector.show())

        # 材料分類輸入區
        row3 = tk.Frame(material_frame)
        row3.pack(fill="x", pady=5)

        tk.Label(row3, text="材料分類:", width=15, anchor='w').pack(side="left")
        self.category_entry = tk.Entry(row3, justify='left')
        self.category_entry.pack(side="left", fill="x", expand=True, padx=5)

        # 建立材料分類選擇器
        self.category_selector = PopupSelector(
            self.main_content,
            lambda: self.odoo_util.get_setup('product_catelog'),
            self.category_entry,
            ["value"],
            "選擇材料分類"
        )
        self.category_entry.bind("<FocusIn>",
                                 lambda e: self.category_selector.show())

        # 加工框架
        process_frame = tk.LabelFrame(form_layout, text="加工", padx=10, pady=10)
        process_frame.pack(fill="x", padx=10, pady=5)

        # 加工流程輸入區
        row4 = tk.Frame(process_frame)
        row4.pack(fill="x", pady=5)

        tk.Label(row4, text="加工流程:", width=15, anchor='w').pack(side="left")
        self.process_entry = tk.Entry(row4, justify='left')
        self.process_entry.pack(side="left", fill="x", expand=True, padx=5)

        # 加工流程選擇器
        process_selector = PopupSelector(
            self.main_content,
            lambda: self.odoo_util.get_setup('operation_flow'),
            self.process_entry,
            ["value"],
            "選擇加工流程"
        )
        self.process_entry.bind("<FocusIn>",
                                lambda e: process_selector.show())
        
        # 表面處理輸入區
        row5 = tk.Frame(process_frame)
        row5.pack(fill="x", pady=5)

        tk.Label(row5, text="表面處理:", width=15, anchor='w').pack(side="left")
        self.surface_entry = tk.Entry(row5, justify='left')
        self.surface_entry.pack(side="left", fill="x", expand=True, padx=5)

        # 表面處理選擇器
        surface_selector = PopupSelector(
            self.main_content,
            lambda: self.odoo_util.get_setup('surface_treatment'),
            self.surface_entry,
            ["value"],
            "選擇表面處理"
        )
        self.surface_entry.bind("<FocusIn>",
                                lambda e: surface_selector.show())


        # 顏色框架
        color_frame = tk.LabelFrame(form_layout, text="顏色", padx=10, pady=10)
        color_frame.pack(fill="x", padx=10, pady=5)

        # 顏色理輸入區
        row6 = tk.Frame(color_frame)
        row6.pack(fill="x", pady=5)

        tk.Label(row6, text="顏色:", width=15, anchor='w').pack(side="left")
        self.color_entry = tk.Entry(row6, justify='left')
        self.color_entry.pack(side="left", fill="x", expand=True, padx=5)

        tk.Label(row6, text="色號:", width=15, anchor='w').pack(side="left")
        self.color_no_entry = tk.Entry(row6, width=8, state='readonly', justify='left')
        self.color_no_entry.pack(side="left", padx=5)

        # 顏色選擇器
        color_selector = PopupSelector(
            self.main_content,
            lambda: self.odoo_util.get_color(self.project_id),
            self.color_entry,
            ["name", "color_no"],
            "選擇顏色",
            additional_target_entries={"color_no": self.color_no_entry}
        )
        self.color_entry.bind("<FocusIn>",
                              lambda e: color_selector.show())

        # 建立按鈕框架並置中
        button_frame = tk.Frame(form_layout)
        button_frame.pack(side="bottom", pady=10)
        button_frame.pack_propagate(False)
        button_frame.configure(width=400, height=40)  # 設定框架寬度

        # 確定按鈕
        self.submit_button = tk.Button(button_frame, text="確定", command=self.submit, width=15)
        self.submit_button.pack(side="left", padx=20, expand=True)

        # 取消按鈕
        self.cancel_button = tk.Button(button_frame, text="取消", command=self.cancel, width=15)
        self.cancel_button.pack(side="left", padx=20, expand=True)

    def clear_main_content(self):
        for widget in self.main_content.winfo_children():
            widget.destroy()

    def submit(self):
        material = self.material_entry.get()
        spec = self.spec_entry.get()
        category = self.category_entry.get()
        process = self.process_entry.get()
        surface = self.surface_entry.get()
        color = self.color_entry.get()
        color_no = self.color_no_entry.get()  # Retrieve color_no

        if not all([material, spec, category, process, surface, color, color_no]):
            # self.show_message("提示", "請填寫所有欄位")
            self.util_log.safe_log_insert("請填寫所有欄位。\n")
            return


        try:
            # 假設所有屬性都在同一個塊中，替換為實際的塊名稱
            active_layout = self.autocad_util.get_active_layout()
            block = self.autocad_util.get_attribute_block(active_layout)
            self.autocad_util.process_pr_no(active_layout)

            if block:
                # 準備屬性值字典
                attr_values = {
                    'product_name': material,
                    'spec': spec,
                    'product_catelog': category,
                    'operation_flow': process,
                    'surface_treatment': surface,
                    'color_name': color,
                    'color_no': color_no
                }

                # 更新每個屬性
                for tag, value in attr_values.items():
                    self.autocad_util.set_attribute_value(block, tag, value)

                # self.show_message("提示", "參數已更新至 AutoCAD。", "info")
                self.util_log.safe_log_insert("參數已更新至 AutoCAD。\n")
                self.clear_main_content()  # 清除內容並重置表單
            else:
                # self.show_message("錯誤", "未找到指定的塊來更新屬性。")
                self.util_log.safe_log_insert("未找到指定的塊來更新屬性。\n")

        except Exception as e:
            # self.show_message("錯誤", f"更新 AutoCAD 屬性時發生錯誤: {str(e)}")
            self.util_log.safe_log_insert(f"更新 AutoCAD 屬性時發生錯誤: {str(e)}\n")

    def cancel(self):
        self.clear_main_content()

    def show_message(self, title, message, msg_type="error"):
        # 确保窗口已更新
        self.root.update_idletasks()
        
        # 创建一个隐藏的 Toplevel 作为父窗口
        top = Toplevel(self.root)
        top.withdraw()  # 隐藏窗口
        top.transient(self.root)  # 设置为对话框
        top.grab_set()  # 捕获所有事件

        # 显示消息
        if msg_type == "info":
            messagebox.showinfo(title, message, parent=top)
        else:
            messagebox.showerror(title, message, parent=top)
        
        # 释放捕获并销毁窗口
        top.grab_release()
        top.destroy()
