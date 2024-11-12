# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox

class PopupSelector:
    def __init__(self, parent, data_source, target_entry, columns, title="選擇一個選項", additional_target_entries=None):
        self.parent = parent
        self.data_source = data_source
        self.target_entry = target_entry
        self.columns = columns
        self.title = title
        self.popup_open = False
        self.filtered_items = []
        self.items = []
        self.additional_target_entries = additional_target_entries or {}

    def show(self):
        if self.popup_open:
            return
        self.popup_open = True
        self.items = self.data_source() or []  # 確保返回空列表而不是 None
        self.filtered_items = self.items
        if not self.filtered_items:
            messagebox.showinfo("提示", "沒有可用的選項")
            self.popup_open = False
            return
        self._create_popup()

    def _create_popup(self):
        self.popup = tk.Toplevel(self.parent)
        self.popup.title(self.title)
        self.popup.geometry("400x300")

        # 更新彈出視窗以獲取正確尺寸
        self.popup.update_idletasks()
        # 將彈出視窗置中
        self._center_popup()

        # 綁定關閉事件
        self.popup.protocol("WM_DELETE_WINDOW", self.on_close)

        self.search_input = tk.Entry(self.popup)
        self.search_input.pack(fill="x")
        self.search_input.bind("<KeyRelease>", self.update_filter)

        self.tree = ttk.Treeview(self.popup, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col.title())
        self.tree.pack(fill="both", expand=True)

        self._update_tree_items()
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

    def _center_popup(self):
        # 更新父視窗和彈出視窗以獲取正確尺寸
        self.parent.update_idletasks()
        self.popup.update_idletasks()

        # 獲取父視窗的位置和尺寸
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()

        # 獲取彈出視窗的尺寸
        popup_width = self.popup.winfo_width()
        popup_height = self.popup.winfo_height()

        # 計算彈出視窗的位置，使其置於父視窗中央
        x = parent_x + (parent_width - popup_width) // 2
        y = parent_y + (parent_height - popup_height) // 2

        # 設定彈出視窗的位置
        self.popup.geometry(f"+{x}+{y}")

    def update_filter(self, event):
        search_text = event.widget.get().lower()
        self.filtered_items = [
            item for item in self.items 
            if any(str(value).lower().find(search_text) != -1 
                  for value in item.values())
        ]
        self._update_tree_items()

    def _update_tree_items(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for item in self.filtered_items:
            values = [item[col] for col in self.columns]
            self.tree.insert("", "end", values=values)

    def on_select(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            values = self.tree.item(selected_item, "values")
            # 將第一列的值插入到 target_entry
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, values[0])
            # 將其他值插入到對應的 Entry 控件
            for i, col in enumerate(self.columns):
                if col in self.additional_target_entries:
                    entry = self.additional_target_entries[col]
                    entry.configure(state='normal')
                    entry.delete(0, tk.END)
                    entry.insert(0, values[i])
                    entry.configure(state='readonly')
            self.popup.destroy()
            self.popup_open = False
            self.target_entry.master.focus_set()

    def on_close(self):
        self.popup.destroy()
        self.popup_open = False

class FormAutoCADParam:
    def __init__(self, main_body, odoo_util):
        self.main_body = main_body
        self.odoo_util = odoo_util
        self.material_selector = None
        self.spec_selector = None
        self.category_selector = None

    def get_parameters_from_odoo(self):
        self.clear_main_body()
        form_layout = tk.Frame(self.main_body)
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
            self.main_body,
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
            self.main_body,
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
            self.main_body,
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
            self.main_body,
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
            self.main_body,
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

        # 顏色選擇器
        color_selector = PopupSelector(
            self.main_body,
            lambda: self.odoo_util.get_color(1),
            self.color_entry,
            ["name"],
            "選擇顏色"
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

    def clear_main_body(self):
        for widget in self.main_body.winfo_children():
            widget.destroy()

    def submit(self):
        material = self.material_entry.get()
        spec = self.spec_entry.get()
        category = self.category_entry.get()
        process = self.process_entry.get()
        surface = self.surface_entry.get()
        color = self.color_entry.get()

        if not all([material, spec, category, process, surface, color]):
            messagebox.showinfo("提示", "請填寫所有欄位")
            return

        messagebox.showinfo("提示", "參數已獲取")

    def cancel(self):
        self.clear_main_body()

