# -*- coding: utf-8 -*-
"""
改進的AutoCAD參數選擇表單
解決原有表單的使用體驗問題
"""

import customtkinter as ctk
from tkinter import messagebox
from typing import Dict, List, Optional, Any
import threading

from ui.ui_theme import theme
from ui.ui_fonts import get_app_font


class EnhancedFormAutoCADParam:
    """改進的AutoCAD參數選擇表單"""
    
    def __init__(self, main_content, odoo_util, autocad_util, log_util, root):
        self.main_content = main_content
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.project_id = self.autocad_util.project_id
        self.job_working_plan_name = self.autocad_util.job_working_plan_name
        self.util_log = log_util
        self.root = root
        
        # 資料快取
        self._data_cache = {}
        self._loading_status = {}
        
        # UI元件
        self.entries = {}
        self.comboboxes = {}
        self.readonly_entries = {}
        
    def get_parameters_from_odoo(self):
        """顯示改進的參數選擇界面"""
        self.clear_main_content()
        
        # 主要容器
        main_frame = ctk.CTkScrollableFrame(
            self.main_content,
            fg_color=theme.get_color('background')
        )
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 標題
        title_label = ctk.CTkLabel(
            main_frame,
            text="🔧 從 Odoo 獲取參數配置",
            font=get_app_font('title'),
            text_color=theme.get_color('text_primary')
        )
        title_label.pack(pady=(0, 20))
        
        # 說明文字
        info_label = ctk.CTkLabel(
            main_frame,
            text="請選擇或輸入所需的參數。可以任意順序填寫，未填寫的欄位將保持空白。",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        info_label.pack(pady=(0, 20))
        
        # 創建各個區塊
        self.create_material_section(main_frame)
        self.create_process_section(main_frame)
        self.create_color_section(main_frame)
        self.create_action_buttons(main_frame)
        
        # 異步載入資料
        self.load_data_async()
    
    def create_material_section(self, parent):
        """創建材料選擇區塊"""
        material_frame = ctk.CTkFrame(parent, fg_color=theme.get_color('surface'))
        material_frame.pack(fill="x", pady=(0, 15))
        
        # 區塊標題
        section_title = ctk.CTkLabel(
            material_frame,
            text="📦 材料配置",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        section_title.pack(pady=(15, 10))
        
        # 材料選擇
        self.create_combobox_field(
            material_frame, 
            "material", 
            "材料:",
            placeholder="選擇或輸入材料名稱...",
            data_key="products",
            readonly_key="material_uom",
            readonly_label="單位"
        )
        
        # 材質選擇
        self.create_combobox_field(
            material_frame,
            "spec",
            "材質:",
            placeholder="選擇或輸入材質...",
            data_key="spec"
        )
        
        # 材料分類選擇
        self.create_combobox_field(
            material_frame,
            "category",
            "材料分類:",
            placeholder="選擇或輸入分類...",
            data_key="product_catelog"
        )
    
    def create_process_section(self, parent):
        """創建加工選擇區塊"""
        process_frame = ctk.CTkFrame(parent, fg_color=theme.get_color('surface'))
        process_frame.pack(fill="x", pady=(0, 15))
        
        section_title = ctk.CTkLabel(
            process_frame,
            text="⚙️ 加工配置",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        section_title.pack(pady=(15, 10))
        
        # 加工流程
        self.create_combobox_field(
            process_frame,
            "process",
            "加工流程:",
            placeholder="選擇或輸入加工流程...",
            data_key="operation_flow"
        )
        
        # 表面處理
        self.create_combobox_field(
            process_frame,
            "surface",
            "表面處理:",
            placeholder="選擇或輸入表面處理方式...",
            data_key="surface_treatment"
        )
    
    def create_color_section(self, parent):
        """創建顏色選擇區塊"""
        color_frame = ctk.CTkFrame(parent, fg_color=theme.get_color('surface'))
        color_frame.pack(fill="x", pady=(0, 15))
        
        section_title = ctk.CTkLabel(
            color_frame,
            text="🎨 顏色配置",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        section_title.pack(pady=(15, 10))
        
        # 顏色選擇
        self.create_combobox_field(
            color_frame,
            "color",
            "顏色:",
            placeholder="選擇或輸入顏色名稱...",
            data_key="colors",
            readonly_key="color_no",
            readonly_label="色號"
        )
    
    def create_combobox_field(self, parent, field_key: str, label: str, 
                             placeholder: str, data_key: str,
                             readonly_key: Optional[str] = None,
                             readonly_label: Optional[str] = None):
        """創建可搜尋的下拉選擇欄位"""
        
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", padx=20, pady=5)
        
        # 標籤
        label_widget = ctk.CTkLabel(
            field_frame,
            text=label,
            font=get_app_font('body'),
            text_color=theme.get_color('text_primary'),
            width=100,
            anchor="w"
        )
        label_widget.pack(side="left", padx=(0, 10))
        
        # 主要輸入欄位 - 使用可編輯的Combobox
        combobox = ctk.CTkComboBox(
            field_frame,
            font=get_app_font('body'),
            command=lambda value, key=field_key: self.on_selection_change(key, value)
        )
        combobox.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # 設置為可編輯，並設置初始提示文字
        combobox.configure(state="normal")
        combobox.set(placeholder)  # 設置提示文字
        
        self.comboboxes[field_key] = {
            'widget': combobox,
            'data_key': data_key,
            'placeholder': placeholder
        }
        
        # 只讀欄位（如單位、色號等）
        if readonly_key and readonly_label:
            readonly_label_widget = ctk.CTkLabel(
                field_frame,
                text=readonly_label + ":",
                font=get_app_font('body'),
                text_color=theme.get_color('text_secondary'),
                width=60,
                anchor="w"
            )
            readonly_label_widget.pack(side="left", padx=(10, 5))
            
            readonly_entry = ctk.CTkEntry(
                field_frame,
                font=get_app_font('body'),
                width=80,
                state="disabled"
            )
            readonly_entry.pack(side="left", padx=(0, 10))
            
            self.readonly_entries[readonly_key] = readonly_entry
    
    def create_action_buttons(self, parent):
        """創建操作按鈕"""
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(fill="x", pady=20)
        
        # 按鈕容器居中
        center_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        center_frame.pack(expand=True)
        
        # 確定按鈕
        submit_btn = ctk.CTkButton(
            center_frame,
            text="✅ 確定並更新到 AutoCAD",
            command=self.submit,
            font=get_app_font('button'),
            height=40,
            width=200,
            fg_color=theme.get_color('success'),
            hover_color=theme.get_color('success_hover')
        )
        submit_btn.pack(side="left", padx=10)
        
        # 重置按鈕
        reset_btn = ctk.CTkButton(
            center_frame,
            text="🔄 重置表單",
            command=self.reset_form,
            font=get_app_font('button'),
            height=40,
            width=120,
            fg_color=theme.get_color('warning'),
            hover_color=theme.get_color('warning_hover')
        )
        reset_btn.pack(side="left", padx=10)
        
        # 取消按鈕
        cancel_btn = ctk.CTkButton(
            center_frame,
            text="❌ 取消",
            command=self.cancel,
            font=get_app_font('button'),
            height=40,
            width=100,
            fg_color=theme.get_color('error'),
            hover_color=theme.get_color('error_hover')
        )
        cancel_btn.pack(side="left", padx=10)
    
    def load_data_async(self):
        """異步載入所有資料"""
        def load_thread():
            self.util_log.safe_log_insert("正在載入參數資料...\n")
            
            # 載入各種資料
            data_loaders = {
                'products': lambda: self.odoo_util.get_product(),
                'spec': lambda: self.odoo_util.get_setup('spec'),
                'product_catelog': lambda: self.odoo_util.get_setup('product_catelog'),
                'operation_flow': lambda: self.odoo_util.get_setup('operation_flow'),
                'surface_treatment': lambda: self.odoo_util.get_setup('surface_treatment'),
                'colors': lambda: self.odoo_util.get_color(self.project_id)
            }
            
            for data_key, loader in data_loaders.items():
                try:
                    self._loading_status[data_key] = "loading"
                    data = loader()
                    self._data_cache[data_key] = data or []
                    self._loading_status[data_key] = "completed"
                    
                    # 更新對應的UI
                    self.root.after(0, lambda dk=data_key: self.update_combobox_data(dk))
                    
                except Exception as e:
                    self._loading_status[data_key] = "error"
                    self.util_log.safe_log_insert(f"載入 {data_key} 資料時發生錯誤: {str(e)}\n")
            
            self.util_log.safe_log_insert("參數資料載入完成\n")
        
        # 在背景線程中載入
        threading.Thread(target=load_thread, daemon=True).start()
    
    def update_combobox_data(self, data_key: str):
        """更新Combobox的選項資料"""
        data = self._data_cache.get(data_key, [])
        
        for field_key, combobox_info in self.comboboxes.items():
            if combobox_info['data_key'] == data_key:
                combobox = combobox_info['widget']
                
                # 準備選項列表
                if data_key == 'products':
                    values = [item.get('name', '') for item in data if item.get('name')]
                elif data_key == 'colors':
                    values = [item.get('name', '') for item in data if item.get('name')]
                else:
                    values = [item.get('value', '') for item in data if item.get('value')]
                
                # 更新Combobox選項
                combobox.configure(values=values)
                
                # 清除提示文字並重新設置（如果有資料的話）
                if values:
                    current_value = combobox.get()
                    if current_value == combobox_info['placeholder']:
                        combobox.set("")  # 清除提示文字
                
                # 如果還沒有資料，顯示載入狀態
                if not values:
                    loading_status = self._loading_status.get(data_key, "")
                    if loading_status == "loading":
                        combobox.set("載入中...")
                    elif loading_status == "error":
                        combobox.set("載入失敗")
    
    def on_selection_change(self, field_key: str, value: str):
        """處理選擇變更事件"""
        if not value:
            return
            
        # 根據欄位類型處理相關資料
        if field_key == "material":
            self.update_material_info(value)
        elif field_key == "color":
            self.update_color_info(value)
    
    def update_material_info(self, material_name: str):
        """更新材料相關資訊"""
        products = self._data_cache.get('products', [])
        for product in products:
            if product.get('name') == material_name:
                uom = product.get('uom', '')
                if 'material_uom' in self.readonly_entries:
                    entry = self.readonly_entries['material_uom']
                    entry.configure(state="normal")
                    entry.delete(0, "end")
                    entry.insert(0, uom)
                    entry.configure(state="disabled")
                break
    
    def update_color_info(self, color_name: str):
        """更新顏色相關資訊"""
        colors = self._data_cache.get('colors', [])
        for color in colors:
            if color.get('name') == color_name:
                color_no = color.get('color_no', '')
                if 'color_no' in self.readonly_entries:
                    entry = self.readonly_entries['color_no']
                    entry.configure(state="normal")
                    entry.delete(0, "end")
                    entry.insert(0, color_no)
                    entry.configure(state="disabled")
                break
    
    def reset_form(self):
        """重置表單"""
        for combobox_info in self.comboboxes.values():
            # 重新設置提示文字
            combobox_info['widget'].set(combobox_info['placeholder'])
            
        for entry in self.readonly_entries.values():
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.configure(state="disabled")
        
        self.util_log.safe_log_insert("表單已重置\n")
    
    def get_form_values(self) -> Dict[str, str]:
        """獲取表單所有值"""
        values = {}
        
        # 獲取主要欄位值
        field_mapping = {
            'material': 'product_name',
            'spec': 'spec',
            'category': 'product_catelog', 
            'process': 'operation_flow',
            'surface': 'surface_treatment',
            'color': 'color_name'
        }
        
        for field_key, attr_key in field_mapping.items():
            if field_key in self.comboboxes:
                value = self.comboboxes[field_key]['widget'].get().strip()
                placeholder = self.comboboxes[field_key]['placeholder']
                # 忽略提示文字
                if value and value != placeholder:
                    values[attr_key] = value
                else:
                    values[attr_key] = ""
        
        # 獲取只讀欄位值
        readonly_mapping = {
            'material_uom': 'uom',
            'color_no': 'color_no'
        }
        
        for field_key, attr_key in readonly_mapping.items():
            if field_key in self.readonly_entries:
                entry = self.readonly_entries[field_key]
                entry.configure(state="normal")
                value = entry.get().strip()
                entry.configure(state="disabled")
                values[attr_key] = value
        
        return values
    
    def submit(self):
        """提交表單"""
        values = self.get_form_values()
        
        # 檢查是否至少填寫了一個欄位
        filled_values = {k: v for k, v in values.items() if v}
        if not filled_values:
            messagebox.showwarning("提示", "請至少填寫一個參數欄位。")
            return
        
        try:
            # 獲取AutoCAD塊
            active_layout = self.autocad_util.get_active_layout()
            block = self.autocad_util.get_attribute_block(active_layout)
            self.autocad_util.process_pr_no(active_layout)
            
            if block:
                # 更新屬性
                for attr_key, value in filled_values.items():
                    if value:  # 只更新有值的屬性
                        self.autocad_util.set_attribute_value(block, attr_key, value)
                
                filled_count = len(filled_values)
                self.util_log.safe_log_insert(f"已更新 {filled_count} 個參數到 AutoCAD\n")
                
                # 顯示成功訊息
                messagebox.showinfo("成功", f"已成功更新 {filled_count} 個參數到 AutoCAD。")
                self.clear_main_content()
                
            else:
                self.util_log.safe_log_insert("未找到指定的塊來更新屬性\n")
                messagebox.showerror("錯誤", "未找到指定的塊來更新屬性。")
                
        except Exception as e:
            error_msg = f"更新 AutoCAD 屬性時發生錯誤: {str(e)}"
            self.util_log.safe_log_insert(error_msg + "\n")
            messagebox.showerror("錯誤", error_msg)
    
    def cancel(self):
        """取消操作"""
        self.clear_main_content()
        self.util_log.safe_log_insert("已取消參數配置\n")
    
    def clear_main_content(self):
        """清除主要內容"""
        for widget in self.main_content.winfo_children():
            widget.destroy()