# -*- coding: utf-8 -*-
import logging
import base64
import requests
import customtkinter as ctk
from tkinter import messagebox
from typing import Optional

from utility.util_odoo import UtilOdoo
from utility.util_autocad import UtilAutoCAD
from utility.util_push_to_boq import UtilPushToBoq
from utility.util_transfer_boq_to_pr import UtilTransferBoqToPr
from utility.util_log import UtilLog
from forms.form_autocad_param import FormAutoCADParam
from forms.form_autocad_param_enhanced import EnhancedFormAutoCADParam
from ui.ui_theme import UITheme, theme
from ui.ui_fonts import get_app_font
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

class ModernFormMain(ctk.CTk):
    """現代化的主表單，使用CustomTkinter"""
    
    def __init__(self, odoo_connection):
        super().__init__()
        
        # 設置主題
        UITheme.setup_theme("system", "blue")
        
        self.odoo_connection = odoo_connection
        self.title("AutoCAD Odoo 整合系統")
        
        # 設置視窗圖示
        try:
            self.iconpath = "icon/odoo_autocad.ico"
            self.iconbitmap(self.iconpath)
        except:
            pass  # 如果圖示檔案不存在，忽略錯誤
        
        # 設置視窗大小和位置
        self.setup_window_geometry()
        
        # 設置視窗樣式
        self.setup_window_style()
        
        # 創建UI (這會創建日誌工具)
        self.create_ui()
        
        # 初始化元件 (現在log_util已經可用)
        self.init_utilities()
        
        # 更新連接狀態
        self.update_connection_status()
    
    def setup_window_geometry(self):
        """設置視窗幾何屬性"""
        # 最小尺寸
        self.minsize(1000, 700)
        
        # 獲取螢幕尺寸
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 預設視窗大小
        window_width = min(1200, int(screen_width * 0.8))
        window_height = min(800, int(screen_height * 0.8))
        
        # 計算置中位置
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    
    def setup_window_style(self):
        """設置視窗樣式"""
        # 設置關閉事件
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 設置視窗圖標(如果可用)
        self.after(100, self.load_window_icon)
    
    def load_window_icon(self):
        """載入視窗圖示"""
        try:
            self.iconbitmap("icon/odoo_autocad.ico")
        except:
            pass
    
    def init_utilities(self):
        """初始化工具類別"""
        # 注意：log_util 已經在 create_ui() -> setup_log_util() 中創建
        
        # 初始化工具類別，現在可以使用 log_util
        self.odoo_util = UtilOdoo(self.odoo_connection, self.log_util)
        self.autocad_util = UtilAutoCAD(self.odoo_util, self.log_util)
        self.push_to_boq_util = UtilPushToBoq(self.odoo_util, self.autocad_util, self.log_util)
        self.transfer_boq_to_pr_util = UtilTransferBoqToPr(self.odoo_util, self.autocad_util, self.log_util)
    
    def create_ui(self):
        """創建使用者介面"""
        # 主要容器
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 頂部標題列
        self.create_top_bar()
        
        # 左側邊欄
        self.create_sidebar()
        
        # 主要內容區域
        self.create_main_content()
        
        # 底部狀態列和日誌
        self.create_bottom_area()
        
        # 設置日誌工具並更新工具類別
        self.setup_log_util()
    
    def create_top_bar(self):
        """創建頂部標題列"""
        self.top_bar = ctk.CTkFrame(
            self, 
            height=theme.get_size('topbar_height'),
            corner_radius=0,
            fg_color=theme.get_color('primary')
        )
        self.top_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        self.top_bar.grid_columnconfigure(1, weight=1)
        
        # 應用程式標題
        title_label = ctk.CTkLabel(
            self.top_bar,
            text="🏢 AutoCAD Odoo 整合系統",
            font=get_app_font('title'),
            text_color="white"
        )
        title_label.grid(row=0, column=0, padx=20, pady=15, sticky="w")
        
        # 連接狀態指示器區域
        self.status_frame = ctk.CTkFrame(self.top_bar, fg_color="transparent")
        self.status_frame.grid(row=0, column=1, padx=20, pady=10, sticky="e")
        
        # Odoo連接狀態
        self.odoo_status_label = ctk.CTkLabel(
            self.status_frame,
            text="🏢 Odoo: 未連接",
            font=get_app_font('small'),
            text_color="white"
        )
        self.odoo_status_label.grid(row=0, column=0, padx=10)
        
        # AutoCAD連接狀態
        self.autocad_status_label = ctk.CTkLabel(
            self.status_frame,
            text="📐 AutoCAD: 未連接",
            font=get_app_font('small'),
            text_color="white"
        )
        self.autocad_status_label.grid(row=0, column=1, padx=10)
    
    def create_sidebar(self):
        """創建左側邊欄"""
        self.sidebar = ctk.CTkFrame(
            self,
            width=theme.get_size('sidebar_width'),
            corner_radius=0,
            fg_color=theme.get_color('surface')
        )
        self.sidebar.grid(row=1, column=0, sticky="nsw", padx=0, pady=0)
        self.sidebar.grid_propagate(False)
        
        # 側邊欄標題
        sidebar_title = ctk.CTkLabel(
            self.sidebar,
            text="功能選單",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        sidebar_title.pack(pady=(20, 10))
        
        # 連接功能區
        self.create_connection_section()
        
        # 分隔線
        separator1 = ctk.CTkFrame(self.sidebar, height=2, fg_color=theme.get_color('border'))
        separator1.pack(fill="x", padx=20, pady=10)
        
        # 主要功能區
        self.create_main_functions_section()
        
        # 分隔線
        separator2 = ctk.CTkFrame(self.sidebar, height=2, fg_color=theme.get_color('border'))
        separator2.pack(fill="x", padx=20, pady=10)
        
        # 工具功能區
        self.create_tools_section()
    
    def create_connection_section(self):
        """創建連接功能區"""
        # 連接區域標題
        conn_label = ctk.CTkLabel(
            self.sidebar,
            text="📡 連接管理",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        conn_label.pack(pady=(0, 10))
        
        # 連接到Odoo按鈕
        self.btn_connect_odoo = ctk.CTkButton(
            self.sidebar,
            text="🏢 連接到 Odoo",
            command=self.connect_odoo,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        self.btn_connect_odoo.pack(fill="x", padx=20, pady=5)
        
        # 連接到AutoCAD按鈕
        self.btn_connect_autocad = ctk.CTkButton(
            self.sidebar,
            text="📐 連接到 AutoCAD",
            command=self.connect_autocad,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        self.btn_connect_autocad.pack(fill="x", padx=20, pady=5)
    
    def create_main_functions_section(self):
        """創建主要功能區"""
        # 主要功能標題
        main_label = ctk.CTkLabel(
            self.sidebar,
            text="⚙️ 主要功能",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        main_label.pack(pady=(0, 10))
        
        # 獲取參數按鈕
        self.btn_get_params = ctk.CTkButton(
            self.sidebar,
            text="📋 從 Odoo 獲取參數",
            command=self.get_parameters_from_odoo,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        self.btn_get_params.pack(fill="x", padx=20, pady=5)
        
        # 推送到BOQ按鈕
        self.btn_push_boq = ctk.CTkButton(
            self.sidebar,
            text="📊 推送到 BOQ",
            command=self.push_to_boq,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        self.btn_push_boq.pack(fill="x", padx=20, pady=5)
        
        # 轉移BOQ到PR按鈕
        self.btn_transfer_pr = ctk.CTkButton(
            self.sidebar,
            text="🔄 轉移 BOQ 到 PR",
            command=self.transfer_boq_to_pr,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        self.btn_transfer_pr.pack(fill="x", padx=20, pady=5)
    
    def create_tools_section(self):
        """創建工具功能區"""
        # 工具標題
        tools_label = ctk.CTkLabel(
            self.sidebar,
            text="🔧 工具",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        tools_label.pack(pady=(0, 10))
        
        # 清除表格ID按鈕
        self.btn_clear_table = ctk.CTkButton(
            self.sidebar,
            text="🗑️ 清除此配置表格ID",
            command=self.clear_table_id,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius'),
            fg_color=theme.get_color('warning'),
            hover_color=theme.get_color('warning_hover')
        )
        self.btn_clear_table.pack(fill="x", padx=20, pady=5)
        
        # 清除所有表格ID按鈕
        self.btn_clear_all_tables = ctk.CTkButton(
            self.sidebar,
            text="🗑️ 清除所有配置表格ID",
            command=self.clear_all_tables_id,
            height=theme.get_size('button_height'),
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius'),
            fg_color=theme.get_color('error'),
            hover_color=theme.get_color('error_hover')
        )
        self.btn_clear_all_tables.pack(fill="x", padx=20, pady=5)
    
    def create_main_content(self):
        """創建主要內容區域"""
        self.main_content = ctk.CTkFrame(
            self,
            corner_radius=theme.get_size('border_radius'),
            fg_color=theme.get_color('background')
        )
        self.main_content.grid(row=1, column=1, sticky="nsew", padx=10, pady=(0, 10))
        
        # 歡迎畫面
        self.create_welcome_screen()
    
    def create_welcome_screen(self):
        """創建歡迎畫面"""
        welcome_frame = ctk.CTkFrame(
            self.main_content,
            fg_color="transparent"
        )
        welcome_frame.pack(expand=True, fill="both")
        
        # 歡迎標題
        welcome_title = ctk.CTkLabel(
            welcome_frame,
            text="歡迎使用 AutoCAD Odoo 整合系統",
            font=get_app_font('title'),
            text_color=theme.get_color('text_primary')
        )
        welcome_title.pack(pady=(50, 20))
        
        # 說明文字
        welcome_text = ctk.CTkLabel(
            welcome_frame,
            text="請先建立 Odoo 和 AutoCAD 的連接，然後開始使用各項功能。",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        welcome_text.pack(pady=(0, 30))
        
        # 快速開始按鈕
        quick_start_frame = ctk.CTkFrame(welcome_frame, fg_color="transparent")
        quick_start_frame.pack()
        
        quick_odoo = ctk.CTkButton(
            quick_start_frame,
            text="🏢 連接 Odoo",
            command=self.connect_odoo,
            height=40,
            width=150,
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        quick_odoo.pack(side="left", padx=10)
        
        quick_autocad = ctk.CTkButton(
            quick_start_frame,
            text="📐 連接 AutoCAD",
            command=self.connect_autocad,
            height=40,
            width=150,
            font=get_app_font('button'),
            corner_radius=theme.get_size('border_radius')
        )
        quick_autocad.pack(side="left", padx=10)
    
    def create_bottom_area(self):
        """創建底部區域（狀態列和日誌）"""
        self.bottom_frame = ctk.CTkFrame(
            self,
            height=200,
            corner_radius=0,
            fg_color=theme.get_color('surface')
        )
        self.bottom_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        self.bottom_frame.grid_propagate(False)
        
        # 日誌標題
        log_title = ctk.CTkLabel(
            self.bottom_frame,
            text="📝 系統日誌",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        log_title.pack(pady=(10, 5))
        
        # 日誌文字區域將由UtilLog管理
        # 這裡只是創建容器
    
    def setup_log_util(self):
        """設置日誌工具"""
        self.log_util = UtilLog(self.bottom_frame)
        
        # 注意：工具類別已經在 init_utilities() 中使用正確的 log_util 初始化
        # 不需要在這裡更新引用
    
    def update_connection_status(self):
        """更新連接狀態顯示"""
        # 更新Odoo狀態
        if hasattr(self, 'odoo_util') and self.odoo_util.connected_odoo():
            self.odoo_status_label.configure(text="🏢 Odoo: ✅ 已連接")
            self.btn_connect_odoo.configure(
                fg_color=theme.get_color('success'),
                hover_color=theme.get_color('success_hover')
            )
        else:
            self.odoo_status_label.configure(text="🏢 Odoo: ❌ 未連接")
            self.btn_connect_odoo.configure(
                fg_color=theme.get_color('primary'),
                hover_color=theme.get_color('primary_dark')
            )
        
        # 更新AutoCAD狀態
        if hasattr(self, 'autocad_util') and self.autocad_util.connected_autocad():
            self.autocad_status_label.configure(text="📐 AutoCAD: ✅ 已連接")
            self.btn_connect_autocad.configure(
                fg_color=theme.get_color('success'),
                hover_color=theme.get_color('success_hover')
            )
        else:
            self.autocad_status_label.configure(text="📐 AutoCAD: ❌ 未連接")
            self.btn_connect_autocad.configure(
                fg_color=theme.get_color('primary'),
                hover_color=theme.get_color('primary_dark')
            )
    
    # === 事件處理方法 ===
    
    def connect_odoo(self):
        """連接到Odoo"""
        self.clear_main_content()
        if self.log_util:
            self.log_util.safe_log_insert("正在連接到 Odoo...\n")
        
        try:
            odoo, requestOptions, token = self.odoo_util.connect_odoo(self.odoo_connection)
            self.update_connection_status()
            if self.log_util:
                self.log_util.safe_log_insert("✅ 與 Odoo 連線成功\n")
        except requests.exceptions.ConnectionError:
            if self.log_util:
                self.log_util.safe_log_insert("❌ 無法與 Odoo 連線，請檢查網路連接或稍後重試\n")
        except Exception as e:
            if self.log_util:
                self.log_util.safe_log_insert(f"❌ 連接 Odoo 時發生錯誤: {str(e)}\n")
    
    def connect_autocad(self):
        """連接到AutoCAD"""
        if self.log_util:
            self.log_util.safe_log_insert("正在連接到 AutoCAD...\n")
        
        self.autocad_util.connect_autocad(self.main_content)
        self.update_connection_status()
        
        if self.autocad_util.connected_autocad():
            if self.log_util:
                self.log_util.safe_log_insert("✅ 與 AutoCAD 連線成功\n")
    
    def get_parameters_from_odoo(self):
        """從Odoo獲取參數 - 使用改進的界面"""
        if not self.autocad_util.project_id:
            self.show_error_message("錯誤", "請先連接 AutoCAD 並確保已獲取專案資料。")
            return
        
        # 使用改進的參數選擇表單
        enhanced_form = EnhancedFormAutoCADParam(
            main_content=self.main_content,
            odoo_util=self.odoo_util,
            autocad_util=self.autocad_util,
            log_util=self.log_util,
            root=self
        )
        enhanced_form.get_parameters_from_odoo()
    
    def push_to_boq(self):
        """推送到BOQ"""
        self.push_to_boq_util.push_to_boq()
    
    def transfer_boq_to_pr(self):
        """轉移BOQ到PR"""
        self.transfer_boq_to_pr_util.transfer_boq_to_pr()
    
    def clear_table_id(self):
        """清除表格ID"""
        self.autocad_util.clear_table_id()
    
    def clear_all_tables_id(self):
        """清除所有表格ID"""
        self.autocad_util.clear_all_tables_id()
    
    def clear_main_content(self):
        """清除主要內容區域"""
        for widget in self.main_content.winfo_children():
            widget.destroy()
        # 重新創建歡迎畫面
        self.create_welcome_screen()
    
    def show_error_message(self, title: str, message: str):
        """顯示錯誤訊息"""
        self.update_idletasks()
        messagebox.showerror(title, message, parent=self)
    
    def show_info_message(self, title: str, message: str):
        """顯示資訊訊息"""
        self.update_idletasks()
        messagebox.showinfo(title, message, parent=self)
    
    def on_closing(self):
        """視窗關閉事件"""
        if self.log_util:
            self.log_util.safe_log_insert("正在關閉應用程式...\n")
        self.destroy()

if __name__ == '__main__':
    # 測試用連接配置
    odoo_connection = {
        'host': 'your_host',
        'db_name': 'your_db_name',
        'url': 'your_url',
        'token': 'your_token'
    }
    
    app = ModernFormMain(odoo_connection=odoo_connection)
    app.mainloop()