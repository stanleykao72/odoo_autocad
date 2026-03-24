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

        # Layout 多選
        self._layout_checkboxes = {}   # {layout_name: ctk.BooleanVar}
        self._active_layout = None
        
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
            text="請選擇或輸入所需的參數。可以任意順序填寫，未填寫的欄位將保持空白。\n💡 提示：在下拉選單中輸入關鍵字即可快速篩選選項！",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        info_label.pack(pady=(0, 20))
        
        # 創建各個區塊
        self.create_layout_selector(main_frame)
        self.create_material_section(main_frame)
        self.create_process_section(main_frame)
        self.create_color_section(main_frame)
        self.create_action_buttons(main_frame)
        
        # 異步載入資料
        self.load_data_async()
    
    def create_layout_selector(self, parent):
        """創建 Layout 多選區塊 — 下拉式可捲動多選面板"""
        layout_frame = ctk.CTkFrame(parent, fg_color=theme.get_color('surface'))
        layout_frame.pack(fill="x", pady=(0, 15))

        section_title = ctk.CTkLabel(
            layout_frame,
            text="📐 選擇要更新的配置 (Layout)",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        section_title.pack(pady=(15, 10))

        # 取得 active layout（從 AutoCAD 即時讀取，確保與實際一致）
        try:
            self._active_layout = self.autocad_util.get_active_layout()
        except Exception:
            self._active_layout = getattr(self.autocad_util, 'layout_name', None)

        # 取得所有 layouts
        all_layouts = []
        try:
            all_layouts = self.autocad_util.get_doc_layouts() or []
        except Exception as e:
            self.util_log.safe_log_insert(f"取得 Layout 列表失敗: {e}\n")

        if not all_layouts:
            no_layout_label = ctk.CTkLabel(
                layout_frame,
                text="（無法取得配置列表，將僅更新目前配置）",
                font=("Microsoft JhengHei UI", 13),
                text_color=theme.get_color('text_secondary')
            )
            no_layout_label.pack(pady=(0, 10))
            return

        # --- 操作列：下拉按鈕 + 全選 / 取消全選 ---
        ctrl_row = ctk.CTkFrame(layout_frame, fg_color="transparent")
        ctrl_row.pack(fill="x", padx=20, pady=(0, 5))

        # 已選摘要標籤（顯示在下拉按鈕上）
        self._layout_summary_label = ctk.CTkLabel(
            ctrl_row,
            text="",
            font=("Microsoft JhengHei UI", 13),
            text_color=theme.get_color('text_primary'),
            anchor="w"
        )
        self._layout_summary_label.pack(side="left", fill="x", expand=True)

        select_all_btn = ctk.CTkButton(
            ctrl_row, text="全選", width=70, height=28,
            font=("Microsoft JhengHei UI", 12),
            fg_color="#546E7A", hover_color="#37474F",
            command=lambda: self._toggle_all_layouts(True)
        )
        select_all_btn.pack(side="right", padx=(5, 0))

        deselect_all_btn = ctk.CTkButton(
            ctrl_row, text="取消全選", width=90, height=28,
            font=("Microsoft JhengHei UI", 12),
            fg_color="#546E7A", hover_color="#37474F",
            command=lambda: self._toggle_all_layouts(False)
        )
        deselect_all_btn.pack(side="right", padx=(5, 0))

        # --- 可捲動 checkbox 列表 ---
        list_height = min(200, max(100, len(all_layouts) * 32))
        scroll_frame = ctk.CTkScrollableFrame(
            layout_frame,
            height=list_height,
            fg_color="#1A2A3A",
            corner_radius=6,
            border_width=1,
            border_color="#4A6A8A"
        )
        scroll_frame.pack(fill="x", padx=20, pady=(0, 10))

        for layout_name in all_layouts:
            is_active = (layout_name == self._active_layout)
            var = ctk.BooleanVar(value=is_active)

            display_text = f"{layout_name}  ★ (目前)" if is_active else layout_name
            font = ("Microsoft JhengHei UI", 14, "bold") if is_active else ("Microsoft JhengHei UI", 14)
            text_color = "#FFD54F" if is_active else "#E0E0E0"

            cb = ctk.CTkCheckBox(
                scroll_frame,
                text=display_text,
                variable=var,
                font=font,
                text_color=text_color,
                hover_color="#455A64",
                border_color="#78909C",
                fg_color="#2E7D32" if is_active else "#1976D2",
                command=self._update_layout_summary,
            )
            cb.pack(anchor="w", padx=10, pady=2)
            self._layout_checkboxes[layout_name] = var

        # 初始摘要
        self._update_layout_summary()

    def _toggle_all_layouts(self, state: bool):
        """全選或取消全選所有 Layout"""
        for var in self._layout_checkboxes.values():
            var.set(state)
        self._update_layout_summary()

    def _update_layout_summary(self):
        """更新已選配置的摘要文字"""
        if not hasattr(self, '_layout_summary_label'):
            return
        selected = self.get_selected_layouts()
        total = len(self._layout_checkboxes)
        if not selected:
            self._layout_summary_label.configure(text="⚠ 未選擇任何配置")
        elif len(selected) == total and total > 0:
            self._layout_summary_label.configure(text=f"已選擇: 全部 ({total} 個配置)")
        else:
            names = ", ".join(selected[:5])
            suffix = f" ...等 {len(selected)} 個" if len(selected) > 5 else ""
            self._layout_summary_label.configure(text=f"已選擇: {names}{suffix}")

    def get_selected_layouts(self) -> list:
        """取得使用者勾選的 Layout 名稱列表"""
        selected = [name for name, var in self._layout_checkboxes.items() if var.get()]
        if not selected:
            # 沒有任何勾選時 fallback 到 active layout
            if self._active_layout:
                return [self._active_layout]
        return selected

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
        
        # 標籤 - 加粗並加大字體
        label_widget = ctk.CTkLabel(
            field_frame,
            text=label,
            font=("Microsoft JhengHei UI", 18, "bold"),  # 進一步加大並加粗
            text_color=theme.get_color('text_primary'),
            width=100,
            anchor="w"
        )
        label_widget.pack(side="left", padx=(0, 10))
        
        # 主要輸入欄位 - 使用可編輯的Combobox，加大字體
        combobox = ctk.CTkComboBox(
            field_frame,
            font=("Microsoft JhengHei UI", 16),  # 進一步加大字體
            command=lambda value, key=field_key: self.on_selection_change(key, value)
        )
        # 設定50%寬度分配
        combobox.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # 設置為可編輯，並設置初始提示文字
        combobox.configure(state="normal")
        combobox.set(placeholder)  # 設置提示文字
        
        # 綁定鍵盤事件進行即時搜尋
        combobox._entry.bind('<KeyRelease>', lambda event, key=field_key: self.on_keyword_search(event, key))
        
        self.comboboxes[field_key] = {
            'widget': combobox,
            'data_key': data_key,
            'placeholder': placeholder,
            'all_values': [],  # 儲存所有選項，用於搜尋
            'filtered_values': []  # 儲存篩選後的選項
        }
        
        # 只讀欄位（如單位、色號等）
        if readonly_key and readonly_label:
            # 創建右側容器，佔50%寬度
            readonly_frame = ctk.CTkFrame(field_frame, fg_color="transparent")
            readonly_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))
            
            readonly_label_widget = ctk.CTkLabel(
                readonly_frame,
                text=readonly_label + ":",
                font=("Microsoft JhengHei UI", 16, "bold"),  # 進一步加大並加粗
                text_color=theme.get_color('text_secondary'),
                width=60,
                anchor="w"
            )
            readonly_label_widget.pack(side="left", padx=(10, 5))
            
            readonly_entry = ctk.CTkEntry(
                readonly_frame,
                font=("Microsoft JhengHei UI", 16),  # 進一步加大字體
                state="disabled"
            )
            readonly_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
            
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
            font=("Microsoft JhengHei UI", 16, "bold"),
            height=50,
            width=220,
            fg_color="#2E7D32",  # 深綠色，更好的對比度
            hover_color="#1B5E20",  # 更深的綠色
            text_color="white"
        )
        submit_btn.pack(side="left", padx=10)
        
        # 重置按鈕
        reset_btn = ctk.CTkButton(
            center_frame,
            text="🔄 重置表單",
            command=self.reset_form,
            font=("Microsoft JhengHei UI", 14, "bold"),
            height=50,
            width=140,
            fg_color="#F57C00",  # 橘色，更好的對比度
            hover_color="#E65100",  # 更深的橘色
            text_color="white"
        )
        reset_btn.pack(side="left", padx=10)
        
        # 取消按鈕
        cancel_btn = ctk.CTkButton(
            center_frame,
            text="❌ 取消",
            command=self.cancel,
            font=("Microsoft JhengHei UI", 14, "bold"),
            height=50,
            width=120,
            fg_color="#D32F2F",  # 深紅色，更好的對比度
            hover_color="#B71C1C",  # 更深的紅色
            text_color="white"
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
                
                # 儲存完整的選項列表以供搜尋使用
                combobox_info['all_values'] = values
                combobox_info['filtered_values'] = values.copy()
                
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
    
    def on_keyword_search(self, event, field_key: str):
        """處理關鍵字搜尋事件"""
        if field_key not in self.comboboxes:
            return
            
        combobox_info = self.comboboxes[field_key]
        combobox = combobox_info['widget']
        keyword = combobox.get().strip().lower()
        
        # 如果是提示文字，不進行搜尋
        if keyword == combobox_info['placeholder'].lower():
            return
            
        # 如果關鍵字為空，顯示所有選項
        if not keyword:
            filtered_values = combobox_info['all_values']
        else:
            # 根據關鍵字篩選選項 - 多種匹配方式
            all_values = combobox_info['all_values']
            filtered_values = []
            
            # 1. 完全匹配（優先級最高）
            exact_matches = [value for value in all_values if value.lower() == keyword]
            
            # 2. 開頭匹配
            starts_with = [value for value in all_values 
                          if value.lower().startswith(keyword) and value.lower() != keyword]
            
            # 3. 包含匹配
            contains = [value for value in all_values 
                       if keyword in value.lower() and not value.lower().startswith(keyword)]
            
            # 按優先級組合結果
            filtered_values = exact_matches + starts_with + contains
        
        # 更新篩選後的選項
        combobox_info['filtered_values'] = filtered_values
        combobox.configure(values=filtered_values)
        
        # 顯示搜尋結果統計
        if keyword and self.util_log:
            total_count = len(combobox_info['all_values'])
            filtered_count = len(filtered_values)
            field_name = {
                'material': '材料',
                'spec': '材質', 
                'category': '材料分類',
                'process': '加工流程',
                'surface': '表面處理',
                'color': '顏色'
            }.get(field_key, field_key)
            
            if filtered_count == 0:
                self.util_log.safe_log_insert(f"🔍 {field_name}搜尋 '{keyword}': 未找到匹配項目\n")
            elif filtered_count == total_count:
                pass  # 顯示全部時不記錄
            else:
                self.util_log.safe_log_insert(f"🔍 {field_name}搜尋 '{keyword}': 找到 {filtered_count} 項匹配結果\n")
        
        # 如果找到匹配項目，自動展開下拉選單
        if filtered_values and len(filtered_values) <= 10 and keyword:  # 只在搜尋時且結果不太多時自動展開
            try:
                combobox._open_dropdown_menu()
            except:
                pass  # 如果展開失敗就忽略
    
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
            # 重置搜尋狀態，顯示所有選項
            if combobox_info['all_values']:
                combobox_info['filtered_values'] = combobox_info['all_values'].copy()
                combobox_info['widget'].configure(values=combobox_info['all_values'])
            
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
        """提交表單 — 支援多 Layout 同時更新"""
        values = self.get_form_values()

        # 檢查是否至少填寫了一個欄位
        filled_values = {k: v for k, v in values.items() if v}
        if not filled_values:
            messagebox.showwarning("提示", "請至少填寫一個參數欄位。")
            return

        # 取得使用者勾選的 Layout 列表
        selected_layouts = self.get_selected_layouts()
        if not selected_layouts:
            messagebox.showwarning("提示", "請至少選擇一個配置 (Layout)。")
            return

        try:
            success_layouts = []
            fail_layouts = []

            for layout_name in selected_layouts:
                try:
                    self.autocad_util.set_block_attributes(filled_values, layout_name)
                    success_layouts.append(layout_name)
                    self.util_log.safe_log_insert(
                        f"✅ 配置 {layout_name}: 已更新 {len(filled_values)} 個參數\n")
                except Exception as e:
                    fail_layouts.append(layout_name)
                    self.util_log.safe_log_insert(
                        f"❌ 配置 {layout_name}: 更新失敗 — {e}\n")

            # 顯示結果
            filled_count = len(filled_values)
            if success_layouts and not fail_layouts:
                layout_list = ", ".join(success_layouts)
                msg = f"已成功更新 {filled_count} 個參數到 {len(success_layouts)} 個配置。\n配置: {layout_list}"
                messagebox.showinfo("成功", msg)
                self.clear_main_content()
            elif success_layouts and fail_layouts:
                msg = (f"部分更新完成:\n"
                       f"✅ 成功: {', '.join(success_layouts)}\n"
                       f"❌ 失敗: {', '.join(fail_layouts)}")
                messagebox.showwarning("部分成功", msg)
            else:
                messagebox.showerror("失敗", f"所有配置更新失敗: {', '.join(fail_layouts)}")

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