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
from utility.util_mcp_sse_manager import MCPSSEManager
from ai_assistant.mcp_server_manager import MCPServerManager
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
        
        # 初始化AI助手相關組件
        self.mcp_server_manager = None  # 將在需要時初始化
        self.mcp_sse_manager = MCPSSEManager(port=8083)  # SSE 伺服器管理器
        self.mcp_sse_manager.set_status_callback(self.on_mcp_sse_status_update)
        
        # 自動啟動 SSE 伺服器以供 Gemini CLI 連接
        self.auto_start_sse_server()
    
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
        self.status_frame.grid_columnconfigure(3, weight=1)  # 讓AI控制區域向右對齊
        
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
        
        # AI助手控制區域
        self.create_ai_control_banner()
        
        # 添加 SSE 狀態指示器
        self.create_sse_status_indicator()
    
    def create_ai_control_banner(self):
        """在頂部banner中創建AI助手控制區域"""
        # MCP 按鈕已移除 - 只保留 SSE 功能
        
        # 保留TCP信息標籤，但使其在banner中更簡潔
        self.tcp_info_label = ctk.CTkLabel(
            self.status_frame,
            text="",  # 初始為空，只在有連接時顯示
            font=("Microsoft JhengHei UI", 9),
            text_color="#E0E0E0"
        )
        self.tcp_info_label.grid(row=1, column=1, columnspan=2, pady=(2, 0), sticky="e")
        
        self.pipe_info_label = None  # 在banner中不顯示pipe信息以節省空間
    
    def create_sse_status_indicator(self):
        """創建 SSE 狀態指示器"""
        # SSE 狀態指示器
        self.sse_status_label = ctk.CTkLabel(
            self.status_frame,
            text="🔴",  # 初始為紅色（未啟動）
            font=("Microsoft JhengHei UI", 14),
            text_color="white"
        )
        self.sse_status_label.grid(row=0, column=4, padx=5)
        
        # SSE 控制按鈕
        self.sse_toggle_button = ctk.CTkButton(
            self.status_frame,
            text="🌊",  # SSE 波浪圖示
            command=self.toggle_sse_server,
            height=28,
            width=28,
            font=("Microsoft JhengHei UI", 12),
            corner_radius=14,
            fg_color="#FF9800",  # 橘色 SSE 按鈕
            hover_color="#F57C00",
            text_color="white"
        )
        self.sse_toggle_button.grid(row=0, column=5, padx=5)
        
        # SSE 信息標籤
        self.sse_info_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=("Microsoft JhengHei UI", 9),
            text_color="#E0E0E0"
        )
        self.sse_info_label.grid(row=1, column=4, columnspan=2, pady=(2, 0), sticky="e")
    
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
            font=("Microsoft JhengHei UI", 18, "bold"),
            text_color="#FFFFFF"  # 純白色，確保對比度
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
            font=("Microsoft JhengHei UI", 14, "bold"),
            text_color="#E0E0E0"  # 淺灰色，確保良好對比度
        )
        conn_label.pack(pady=(0, 10))
        
        # 連接到Odoo按鈕
        self.btn_connect_odoo = ctk.CTkButton(
            self.sidebar,
            text="🏢 連接到 Odoo",
            command=self.connect_odoo,
            height=40,
            font=("Microsoft JhengHei UI", 14, "bold"),
            corner_radius=8,
            fg_color="#1976D2",  # 藍色，更好的對比度
            hover_color="#0D47A1",  # 更深的藍色
            text_color="white"
        )
        self.btn_connect_odoo.pack(fill="x", padx=20, pady=5)
        
        # 連接到AutoCAD按鈕
        self.btn_connect_autocad = ctk.CTkButton(
            self.sidebar,
            text="📐 連接到 AutoCAD",
            command=self.connect_autocad,
            height=40,
            font=("Microsoft JhengHei UI", 14, "bold"),
            corner_radius=8,
            fg_color="#7B1FA2",  # 紫色，更好的對比度
            hover_color="#4A148C",  # 更深的紫色
            text_color="white"
        )
        self.btn_connect_autocad.pack(fill="x", padx=20, pady=5)
    
    def create_main_functions_section(self):
        """創建主要功能區"""
        # 主要功能標題
        main_label = ctk.CTkLabel(
            self.sidebar,
            text="⚙️ 主要功能",
            font=("Microsoft JhengHei UI", 14, "bold"),
            text_color="#E0E0E0"  # 淺灰色，確保良好對比度
        )
        main_label.pack(pady=(0, 10))
        
        # 獲取參數按鈕
        self.btn_get_params = ctk.CTkButton(
            self.sidebar,
            text="📋 從 Odoo 獲取參數",
            command=self.get_parameters_from_odoo,
            height=40,
            font=("Microsoft JhengHei UI", 14, "bold"),
            corner_radius=8,
            fg_color="#388E3C",  # 綠色，更好的對比度
            hover_color="#1B5E20",  # 更深的綠色
            text_color="white"
        )
        self.btn_get_params.pack(fill="x", padx=20, pady=5)
        
        # 推送到BOQ按鈕
        self.btn_push_boq = ctk.CTkButton(
            self.sidebar,
            text="📊 推送到 BOQ",
            command=self.push_to_boq,
            height=40,
            font=("Microsoft JhengHei UI", 14, "bold"),
            corner_radius=8,
            fg_color="#F57C00",  # 橘色，更好的對比度
            hover_color="#E65100",  # 更深的橘色
            text_color="white"
        )
        self.btn_push_boq.pack(fill="x", padx=20, pady=5)
        
        # 轉移BOQ到PR按鈕
        self.btn_transfer_pr = ctk.CTkButton(
            self.sidebar,
            text="🔄 轉移 BOQ 到 PR",
            command=self.transfer_boq_to_pr,
            height=40,
            font=("Microsoft JhengHei UI", 14, "bold"),
            corner_radius=8,
            fg_color="#5D4037",  # 棕色，更好的對比度
            hover_color="#3E2723",  # 更深的棕色
            text_color="white"
        )
        self.btn_transfer_pr.pack(fill="x", padx=20, pady=5)
    
    def create_tools_section(self):
        """創建工具功能區"""
        # 工具標題
        tools_label = ctk.CTkLabel(
            self.sidebar,
            text="🔧 工具",
            font=("Microsoft JhengHei UI", 14, "bold"),
            text_color="#E0E0E0"  # 淺灰色，確保良好對比度
        )
        tools_label.pack(pady=(0, 10))
        
        # 清除表格ID按鈕
        self.btn_clear_table = ctk.CTkButton(
            self.sidebar,
            text="🗑️ 清除此配置表格ID",
            command=self.clear_table_id,
            height=40,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#FF8F00",  # 警告橘色，更好的對比度
            hover_color="#E65100",  # 更深的橘色
            text_color="white"
        )
        self.btn_clear_table.pack(fill="x", padx=20, pady=5)
        
        # 清除所有表格ID按鈕
        self.btn_clear_all_tables = ctk.CTkButton(
            self.sidebar,
            text="🗑️ 清除所有配置表格ID",
            command=self.clear_all_tables_id,
            height=40,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#D32F2F",  # 危險紅色，更好的對比度
            hover_color="#B71C1C",  # 更深的紅色
            text_color="white"
        )
        self.btn_clear_all_tables.pack(fill="x", padx=20, pady=5)
        
        # 分隔線
        separator3 = ctk.CTkFrame(self.sidebar, height=2, fg_color=theme.get_color('border'))
        separator3.pack(fill="x", padx=20, pady=10)
        
        # AI助手控制按鈕
        self.btn_sse_control = ctk.CTkButton(
            self.sidebar,
            text="🌊 SSE 伺服器控制",
            command=self.create_sse_control_panel,
            height=40,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#FF9800",  # 橘色，與頂部 SSE 按鈕一致
            hover_color="#F57C00",
            text_color="white"
        )
        self.btn_sse_control.pack(fill="x", padx=20, pady=5)
    
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
                fg_color="#2E7D32",  # 深綠色表示已連接
                hover_color="#1B5E20"
            )
        else:
            self.odoo_status_label.configure(text="🏢 Odoo: ❌ 未連接")
            self.btn_connect_odoo.configure(
                fg_color="#1976D2",  # 原藍色表示未連接
                hover_color="#0D47A1"
            )
        
        # 更新AutoCAD狀態
        if hasattr(self, 'autocad_util') and self.autocad_util.connected_autocad():
            self.autocad_status_label.configure(text="📐 AutoCAD: ✅ 已連接")
            self.btn_connect_autocad.configure(
                fg_color="#2E7D32",  # 深綠色表示已連接
                hover_color="#1B5E20"
            )
        else:
            self.autocad_status_label.configure(text="📐 AutoCAD: ❌ 未連接")
            self.btn_connect_autocad.configure(
                fg_color="#7B1FA2",  # 原紫色表示未連接
                hover_color="#4A148C"
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
    
    def initialize_mcp_server_manager(self):
        """初始化MCP服務管理器"""
        try:
            if self.mcp_server_manager is None:
                self.mcp_server_manager = MCPServerManager(
                    autocad_util=self.autocad_util,
                    odoo_util=self.odoo_util,
                    log_util=self.log_util
                )
                self.log_util.safe_log_insert("MCP服務管理器初始化完成\n")
        except Exception as e:
            self.log_util.safe_log_insert(f"MCP服務管理器初始化失敗: {e}\n")
            messagebox.showerror("錯誤", f"無法初始化AI助手服務管理器：{e}")
    
    # toggle_mcp_server 方法已移除 - MCP 按鈕已從 UI 中移除
    
    # update_mcp_status_display 方法已移除 - MCP 按鈕已從 UI 中移除
    
    def toggle_sse_server(self):
        """切換 SSE 伺服器狀態"""
        self.log_util.safe_log_insert(f"[SSE GUI] toggle_sse_server 被呼叫\n")
        try:
            current_status = self.mcp_sse_manager.is_running
            self.log_util.safe_log_insert(f"[SSE GUI] 當前 SSE 伺服器狀態: {current_status}\n")
            
            if current_status:
                # 停止 SSE 伺服器
                self.log_util.safe_log_insert("[SSE GUI] 準備停止 SSE 伺服器\n")
                stop_result = self.mcp_sse_manager.stop_server()
                self.log_util.safe_log_insert(f"[SSE GUI] 停止 SSE 伺服器結果: {stop_result}\n")
                if stop_result:
                    self.log_util.safe_log_insert("[SSE GUI] ✅ SSE 伺服器已成功停止\n")
                else:
                    self.log_util.safe_log_insert("[SSE GUI] ❌ SSE 伺服器停止失敗\n")
            else:
                # 啟動 SSE 伺服器
                self.log_util.safe_log_insert("[SSE GUI] 準備啟動 SSE 伺服器（端口: 8083）\n")
                import threading
                thread = threading.Thread(target=self._start_sse_server_async, daemon=True)
                self.log_util.safe_log_insert(f"[SSE GUI] 創建啟動執行緒: {thread.name}\n")
                thread.start()
                self.log_util.safe_log_insert("[SSE GUI] 啟動執行緒已開始\n")
                
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.log_util.safe_log_insert(f"[SSE GUI] ❌ 切換 SSE 伺服器狀態發生異常: {e}\n")
            self.log_util.safe_log_insert(f"[SSE GUI] 錯誤追蹤:\n{error_trace}\n")
            messagebox.showerror("錯誤", f"無法切換 SSE 伺服器狀態：{e}")
    
    def auto_start_sse_server(self):
        """自動啟動 SSE 伺服器以供 Gemini CLI 連接"""
        try:
            # 在應用程式啟動時自動啟動 SSE 伺服器
            # 這解決了 Gemini CLI 在 SSE 伺服器啟動之前就嘗試連接的時機問題
            if hasattr(self, 'log_util') and self.log_util:
                self.log_util.safe_log_insert("[SSE GUI] 開始自動啟動 SSE 伺服器流程\n")
                self.log_util.safe_log_insert(f"[SSE GUI] 目標端口: {getattr(self.mcp_sse_manager, 'port', '未知')}\n")
            
            import threading
            thread = threading.Thread(target=self._start_sse_server_async, daemon=True)
            if hasattr(self, 'log_util') and self.log_util:
                self.log_util.safe_log_insert(f"[SSE GUI] 創建自動啟動執行緒: {thread.name}\n")
            thread.start()
            
            if hasattr(self, 'log_util') and self.log_util:
                self.log_util.safe_log_insert("[SSE GUI] 自動啟動執行緒已開始執行\n")
            
        except Exception as e:
            # 如果自動啟動失敗，記錄錯誤但不阻止應用程式啟動
            import traceback
            error_trace = traceback.format_exc()
            if hasattr(self, 'log_util') and self.log_util:
                self.log_util.safe_log_insert(f"[SSE GUI] ❌ 自動啟動 SSE 伺服器失敗: {e}\n")
                self.log_util.safe_log_insert(f"[SSE GUI] 錯誤追蹤:\n{error_trace}\n")
            else:
                print(f"[SSE GUI] 自動啟動 SSE 伺服器失敗: {e}")
                print(f"[SSE GUI] 錯誤追蹤:\n{error_trace}")
    
    def _start_sse_server_async(self):
        """異步啟動 SSE 伺服器"""
        import threading
        thread_name = threading.current_thread().name
        
        # 使用 after 確保日誌在主執行緒中記錄
        self.after(0, lambda: self.log_util.safe_log_insert(f"[SSE GUI] _start_sse_server_async 開始執行 (執行緒: {thread_name})\n"))
        
        try:
            self.after(0, lambda: self.log_util.safe_log_insert("[SSE GUI] 呼叫 mcp_sse_manager.start_server()\n"))
            success = self.mcp_sse_manager.start_server()
            
            self.after(0, lambda: self.log_util.safe_log_insert(f"[SSE GUI] start_server() 回傳結果: {success}\n"))
            
            if success:
                self.after(0, lambda: self.log_util.safe_log_insert("[SSE GUI] ✅ SSE 伺服器啟動成功\n"))
                # 驗證伺服器狀態
                final_status = self.mcp_sse_manager.is_running
                self.after(0, lambda: self.log_util.safe_log_insert(f"[SSE GUI] 最終伺服器狀態: {final_status}\n"))
            else:
                self.after(0, lambda: self.log_util.safe_log_insert("[SSE GUI] ❌ SSE 伺服器啟動失敗\n"))
                
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.after(0, lambda: self.log_util.safe_log_insert(f"[SSE GUI] ❌ _start_sse_server_async 發生異常: {e}\n"))
            self.after(0, lambda: self.log_util.safe_log_insert(f"[SSE GUI] 錯誤追蹤:\n{error_trace}\n"))
    
    def on_mcp_sse_status_update(self, message: str, is_running: bool):
        """SSE 狀態更新回調"""
        def update_ui():
            self.log_util.safe_log_insert(f"[SSE GUI] 收到狀態更新回調: message='{message}', is_running={is_running}\n")
            
            # 更新 SSE 狀態顯示
            if is_running:
                self.log_util.safe_log_insert("[SSE GUI] 更新 UI 為運行狀態\n")
                self.sse_status_label.configure(text="🟢")  # 綠色表示運行中
                self.sse_toggle_button.configure(text="⏹️")  # 停止圖示
                self.sse_info_label.configure(text=f"SSE: :{self.mcp_sse_manager.port}")
            else:
                self.log_util.safe_log_insert("[SSE GUI] 更新 UI 為停止狀態\n")
                self.sse_status_label.configure(text="🔴")  # 紅色表示停止
                self.sse_toggle_button.configure(text="🌊")  # 啟動圖示
                self.sse_info_label.configure(text="")
            
            # 記錄狀態消息
            self.log_util.safe_log_insert(f"[SSE 狀態] {message}\n")
            
            # 更新面板狀態（如果面板已打開）
            try:
                self.update_sse_panel_status()
                self.log_util.safe_log_insert("[SSE GUI] 面板狀態已更新\n")
            except Exception as e:
                self.log_util.safe_log_insert(f"[SSE GUI] 更新面板狀態失敗: {e}\n")
        
        self.after(0, update_ui)
    
    def show_sse_status(self):
        """顯示 SSE 伺服器詳細狀態"""
        status = self.mcp_sse_manager.get_server_status()
        
        status_text = f"""
SSE 伺服器狀態:
運行狀態: {'🟢 運行中' if status['is_running'] else '🔴 已停止'}
端口: {status['port']}
模式: 直接整合 MCPSSEServer
健康檢查: {'✅ 正常' if status['health_check'] else '❌ 異常'}
        """
        
        # 添加伺服器資訊（如果可用）
        if 'server_name' in status:
            status_text += f"\n伺服器名稱: {status['server_name']}"
            status_text += f"\n版本: {status['server_version']}"
            status_text += f"\n活動連接: {status['active_connections']}"
        
        messagebox.showinfo("SSE 伺服器狀態", status_text)
    
    def test_sse_connection(self):
        """測試 SSE 連接"""
        self.log_util.safe_log_insert("[SSE GUI] 開始測試 SSE 連接\n")
        
        try:
            # 記錄測試前的狀態
            server_status = self.mcp_sse_manager.get_server_status()
            self.log_util.safe_log_insert(f"[SSE GUI] 測試前伺服器狀態: {server_status}\n")
            
            # 執行連接測試
            self.log_util.safe_log_insert("[SSE GUI] 呼叫 test_mcp_connection()\n")
            result = self.mcp_sse_manager.test_mcp_connection()
            self.log_util.safe_log_insert(f"[SSE GUI] 測試結果: {result}\n")
            
            if result["success"]:
                test_result = result.get('test_result', '未知')
                message = f"✅ SSE 連接成功\n工具數量: {result['tools_count']}\n可用工具: {', '.join(result['tools'])}\n\n工具測試結果:\n{test_result}"
                self.log_util.safe_log_insert("[SSE GUI] ✅ SSE 連接測試成功\n")
                messagebox.showinfo("SSE 連接測試", message)
            else:
                error_msg = result.get('error', '未知錯誤')
                self.log_util.safe_log_insert(f"[SSE GUI] ❌ SSE 連接測試失敗: {error_msg}\n")
                messagebox.showerror("SSE 連接測試", f"❌ SSE 連接失敗\n錯誤: {error_msg}")
                
        except Exception as e:
            import traceback
            error_trace = traceback.format_exc()
            self.log_util.safe_log_insert(f"[SSE GUI] ❌ 測試 SSE 連接時發生異常: {e}\n")
            self.log_util.safe_log_insert(f"[SSE GUI] 錯誤追蹤:\n{error_trace}\n")
            messagebox.showerror("SSE 連接測試", f"❌ 測試過程發生錯誤：{e}")
    
    def create_sse_control_panel(self):
        """創建 SSE 控制面板"""
        # 清除主要內容
        self.clear_main_content()
        
        # 創建 SSE 控制面板
        panel_frame = ctk.CTkFrame(self.main_content)
        panel_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 標題
        title_label = ctk.CTkLabel(
            panel_frame,
            text="🌊 SSE 伺服器控制面板",
            font=get_app_font('title'),
            text_color=theme.get_color('text_primary')
        )
        title_label.pack(pady=(20, 10))
        
        # 狀態顯示
        status_frame = ctk.CTkFrame(panel_frame)
        status_frame.pack(fill="x", padx=20, pady=10)
        
        self.sse_panel_status_label = ctk.CTkLabel(
            status_frame,
            text="狀態: 未知",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary')
        )
        self.sse_panel_status_label.pack(pady=10)
        
        # 控制按鈕
        button_frame = ctk.CTkFrame(panel_frame)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        # 啟動/停止按鈕
        self.sse_panel_toggle_button = ctk.CTkButton(
            button_frame,
            text="🚀 啟動 SSE 伺服器",
            command=self.toggle_sse_server,
            height=40,
            font=get_app_font('button'),
            corner_radius=8,
            fg_color="#4CAF50",  # 綠色啟動按鈕
            hover_color="#388E3C"
        )
        self.sse_panel_toggle_button.pack(side="left", padx=5)
        
        # 測試連接按鈕
        test_button = ctk.CTkButton(
            button_frame,
            text="🧪 測試連接",
            command=self.test_sse_connection,
            height=40,
            font=get_app_font('button'),
            corner_radius=8,
            fg_color="#2196F3",  # 藍色測試按鈕
            hover_color="#1976D2"
        )
        test_button.pack(side="left", padx=5)
        
        # 查看狀態按鈕
        status_button = ctk.CTkButton(
            button_frame,
            text="📊 查看狀態",
            command=self.show_sse_status,
            height=40,
            font=get_app_font('button'),
            corner_radius=8,
            fg_color="#9C27B0",  # 紫色狀態按鈕
            hover_color="#7B1FA2"
        )
        status_button.pack(side="left", padx=5)
        
        # 配置信息
        config_frame = ctk.CTkFrame(panel_frame)
        config_frame.pack(fill="x", padx=20, pady=10)
        
        config_title = ctk.CTkLabel(
            config_frame,
            text="⚙️ 配置信息",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        config_title.pack(pady=(10, 5))
        
        # 配置詳情
        config_details = ctk.CTkTextbox(
            config_frame,
            height=100,
            font=get_app_font('body')
        )
        config_details.pack(fill="x", padx=10, pady=5)
        
        # 插入配置信息
        config_text = f"""端口: {self.mcp_sse_manager.port}
整合模式: 直接整合 MCPSSEServer 類別
Gemini CLI 配置:
{{
  "autocad-odoo-sse": {{
    "url": "http://localhost:{self.mcp_sse_manager.port}/sse",
    "timeout": 30000,
    "description": "AutoCAD-Odoo Integration with SSE transport"
  }}
}}"""
        config_details.insert("0.0", config_text)
        config_details.configure(state="disabled")
        
        # 使用說明
        help_frame = ctk.CTkFrame(panel_frame)
        help_frame.pack(fill="x", padx=20, pady=10)
        
        help_title = ctk.CTkLabel(
            help_frame,
            text="💡 使用說明",
            font=get_app_font('heading'),
            text_color=theme.get_color('text_primary')
        )
        help_title.pack(pady=(10, 5))
        
        help_text = ctk.CTkLabel(
            help_frame,
            text="1. 點擊 '🚀 啟動 SSE 伺服器' 來啟動伺服器\n2. 伺服器啟動後，可以在 Gemini CLI 中使用 SSE 模式\n3. 使用 '🧪 測試連接' 來驗證伺服器是否正常運行\n4. 查看頂部橫幅的 SSE 狀態指示器瞭解即時狀態",
            font=get_app_font('body'),
            text_color=theme.get_color('text_secondary'),
            justify="left"
        )
        help_text.pack(padx=10, pady=5)
        
        # 更新面板狀態
        self.update_sse_panel_status()
    
    def update_sse_panel_status(self):
        """更新 SSE 面板狀態"""
        if hasattr(self, 'sse_panel_status_label'):
            status = self.mcp_sse_manager.get_server_status()
            if status['is_running']:
                self.sse_panel_status_label.configure(text="狀態: 🟢 運行中")
                if hasattr(self, 'sse_panel_toggle_button'):
                    self.sse_panel_toggle_button.configure(
                        text="⏹️ 停止 SSE 伺服器",
                        fg_color="#f44336",  # 紅色停止按鈕
                        hover_color="#d32f2f"
                    )
            else:
                self.sse_panel_status_label.configure(text="狀態: 🔴 已停止")
                if hasattr(self, 'sse_panel_toggle_button'):
                    self.sse_panel_toggle_button.configure(
                        text="🚀 啟動 SSE 伺服器",
                        fg_color="#4CAF50",  # 綠色啟動按鈕
                        hover_color="#388E3C"
                    )
    
    def on_closing(self):
        """視窗關閉事件"""
        # 在關閉應用程式前停止MCP服務
        try:
            if self.mcp_server_manager and self.mcp_server_manager.is_running():
                self.mcp_server_manager.stop_all_servers()
                self.log_util.safe_log_insert("AI助手服務已停止\n")
        except Exception as e:
            self.log_util.safe_log_insert(f"停止AI助手服務失敗: {e}\n")
        
        # 停止 SSE 伺服器
        try:
            if self.mcp_sse_manager and self.mcp_sse_manager.is_running:
                self.mcp_sse_manager.cleanup()
                self.log_util.safe_log_insert("SSE 伺服器已停止\n")
        except Exception as e:
            self.log_util.safe_log_insert(f"停止 SSE 伺服器失敗: {e}\n")
        
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