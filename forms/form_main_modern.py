# -*- coding: utf-8 -*-
import logging
import os
import sys
import base64
import queue
import threading
import requests
import customtkinter as ctk
from tkinter import messagebox
from typing import Optional

from utility.util_odoo import UtilOdoo
from utility.util_autocad import UtilAutoCAD
from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
from utility.util_push_to_boq import UtilPushToBoq
from utility.util_transfer_boq_to_pr import UtilTransferBoqToPr
from utility.util_log import UtilLog
from utility.util_mcp_manager import MCPManager
from utility.util_gui_proxy import setup_gui_proxy_handlers, get_gui_proxy
from forms.form_autocad_param import FormAutoCADParam
from forms.form_autocad_param_enhanced import EnhancedFormAutoCADParam
from ui.ui_theme import UITheme, theme
from ui.ui_fonts import get_app_font
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

from version import APP_VERSION

class ModernFormMain(ctk.CTk):
    """現代化的主表單，使用CustomTkinter"""

    def __init__(self, odoo_connection, autocad_mode="com"):
        super().__init__()

        # 設置主題
        UITheme.setup_theme("system", "blue")

        self.odoo_connection = odoo_connection
        self.autocad_mode = autocad_mode
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

        # 初始化工具類別
        self.odoo_util = UtilOdoo(self.odoo_connection, self.log_util)

        # 使用 Dispatcher 統一 COM/IPC 介面
        self.autocad_dispatcher = UtilAutoCADDispatcher(
            self.odoo_util, self.log_util, mode=self.autocad_mode
        )
        # 保持向後相容：autocad_util 指向 dispatcher
        self.autocad_util = self.autocad_dispatcher

        self.push_to_boq_util = UtilPushToBoq(self.odoo_util, self.autocad_dispatcher, self.log_util)
        self.transfer_boq_to_pr_util = UtilTransferBoqToPr(self.odoo_util, self.autocad_dispatcher, self.log_util)

        # MCP Manager (取代 MCPSSEManager)
        self.mcp_manager = MCPManager(
            autocad_dispatcher=self.autocad_dispatcher,
            odoo_util=self.odoo_util,
            transport="streamable-http",
            port=8084
        )
        self.mcp_manager.set_status_callback(self.on_mcp_status_update)

        # GUI代理系統：僅 COM 模式需要（解決COM線程問題）
        if self.autocad_mode == "com":
            self.gui_proxy = setup_gui_proxy_handlers(
                self.autocad_dispatcher.active_backend, self.log_util
            )
            self.log_util.safe_log_insert("[GUI] GUI代理系統已初始化 (COM模式)\n")
            self.start_gui_proxy_processing()
        else:
            self.gui_proxy = None
            self.log_util.safe_log_insert("[GUI] IPC模式 — GUI代理系統已跳過\n")

        # 自動啟動 MCP 伺服器
        self.auto_start_mcp_server()
    
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
            text=f"🏢 AutoCAD Odoo 整合系統  v{APP_VERSION}",
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
        """創建左側邊欄（可捲動）"""
        # 外層固定容器
        sidebar_outer = ctk.CTkFrame(
            self,
            width=theme.get_size('sidebar_width'),
            corner_radius=0,
            fg_color=theme.get_color('surface')
        )
        sidebar_outer.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        sidebar_outer.grid_propagate(False)

        # 可捲動的內層
        self.sidebar = ctk.CTkScrollableFrame(
            sidebar_outer,
            corner_radius=0,
            fg_color=theme.get_color('surface'),
            scrollbar_button_color=theme.get_color('border'),
            scrollbar_button_hover_color="#555555"
        )
        self.sidebar.pack(fill="both", expand=True)

        # 側邊欄標題
        sidebar_title = ctk.CTkLabel(
            self.sidebar,
            text="功能選單",
            font=("Microsoft JhengHei UI", 16, "bold"),
            text_color="#FFFFFF"
        )
        sidebar_title.pack(pady=(12, 6))

        # 連接功能區
        self.create_connection_section()

        # 分隔線
        self._sidebar_sep()

        # 圖面資訊區
        self.create_drawing_info_section()

        # 分隔線
        self._sidebar_sep()

        # 主要功能區
        self.create_main_functions_section()

        # 分隔線
        self._sidebar_sep()

        # 工具功能區
        self.create_tools_section()

    def _sidebar_sep(self):
        """側邊欄分隔線"""
        ctk.CTkFrame(self.sidebar, height=1, fg_color=theme.get_color('border')).pack(
            fill="x", padx=15, pady=6
        )
    
    def create_connection_section(self):
        """創建連接功能區"""
        conn_label = ctk.CTkLabel(
            self.sidebar,
            text="📡 連接管理",
            font=("Microsoft JhengHei UI", 13, "bold"),
            text_color="#E0E0E0"
        )
        conn_label.pack(pady=(0, 4))

        self.btn_connect_odoo = ctk.CTkButton(
            self.sidebar,
            text="🏢 連接到 Odoo",
            command=self.connect_odoo,
            height=34,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#1976D2",
            hover_color="#0D47A1",
            text_color="white"
        )
        self.btn_connect_odoo.pack(fill="x", padx=15, pady=3)

        self.btn_connect_autocad = ctk.CTkButton(
            self.sidebar,
            text="📐 連接到 AutoCAD",
            command=self.connect_autocad,
            height=34,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#7B1FA2",
            hover_color="#4A148C",
            text_color="white"
        )
        self.btn_connect_autocad.pack(fill="x", padx=15, pady=3)

        # AutoCAD 模式切換
        mode_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        mode_frame.pack(fill="x", padx=15, pady=(2, 3))

        mode_label = ctk.CTkLabel(
            mode_frame,
            text="模式:",
            font=("Microsoft JhengHei UI", 11),
            text_color="#B0B0B0"
        )
        mode_label.pack(side="left", padx=(0, 5))

        self.mode_var = ctk.StringVar(value=self.autocad_mode.upper())
        self.mode_switch = ctk.CTkSegmentedButton(
            mode_frame,
            values=["COM", "IPC"],
            variable=self.mode_var,
            command=self._on_mode_switch,
            font=("Microsoft JhengHei UI", 11),
            height=28,
            corner_radius=6
        )
        self.mode_switch.pack(side="left", fill="x", expand=True)

    def _on_mode_switch(self, selected):
        """處理 COM/IPC 模式切換"""
        new_mode = selected.lower()
        if new_mode == self.autocad_mode:
            return

        old_mode = self.autocad_mode
        self.autocad_mode = new_mode

        # 切換 dispatcher 模式
        self.autocad_dispatcher.switch_mode(new_mode)

        # COM 模式需要 GUI Proxy，IPC 不需要
        if new_mode == "com" and self.gui_proxy is None:
            from utility.util_gui_proxy import setup_gui_proxy_handlers
            self.gui_proxy = setup_gui_proxy_handlers(
                self.autocad_dispatcher.active_backend, self.log_util
            )
            self.start_gui_proxy_processing()
            self.log_util.safe_log_insert("[GUI] COM 模式 — GUI代理系統已啟動\n")
        elif new_mode == "ipc" and self.gui_proxy is not None:
            self.gui_proxy = None
            self.log_util.safe_log_insert("[GUI] IPC 模式 — GUI代理系統已停用\n")

        # 更新 MCP Manager
        if hasattr(self, 'mcp_manager'):
            self.mcp_manager.update_autocad_dispatcher(self.autocad_dispatcher)

        self.log_util.safe_log_insert(
            f"[Mode] {old_mode.upper()} -> {new_mode.upper()}\n"
        )

        # 更新連接按鈕文字
        self._update_autocad_button_label()

        # 自動連接新模式的 AutoCAD（包含讀取 pr_no 等初始化）
        self.connect_autocad()
        self.update_connection_status()

    def _update_autocad_button_label(self):
        """根據模式更新 AutoCAD 按鈕文字"""
        if self.autocad_mode == "ipc":
            self.btn_connect_autocad.configure(text="📐 連接到 AutoCAD LT")
        else:
            self.btn_connect_autocad.configure(text="📐 連接到 AutoCAD")

    def create_drawing_info_section(self):
        """創建圖面資訊區 — 顯示目前配置/專案/PR"""
        info_label = ctk.CTkLabel(
            self.sidebar,
            text="📄 圖面資訊",
            font=("Microsoft JhengHei UI", 13, "bold"),
            text_color="#E0E0E0"
        )
        info_label.pack(pady=(0, 3))

        # 帶邊框底色的 card，更醒目
        self.info_card = ctk.CTkFrame(
            self.sidebar,
            fg_color="#1E3A5F",
            corner_radius=6,
            border_width=1,
            border_color="#4A6A8A"
        )
        self.info_card.pack(fill="x", padx=10, pady=2)

        dim = "#90A4AE"  # 未連接的預設色（比 #78909C 亮）

        # 目前配置 (Layout)
        self.lbl_layout = ctk.CTkLabel(
            self.info_card,
            text="配置 (Layout): --",
            font=("Microsoft JhengHei UI", 12),
            text_color=dim,
            anchor="w"
        )
        self.lbl_layout.pack(fill="x", padx=8, pady=(5, 1))

        # 請購單號 (PR No)
        self.lbl_pr_no = ctk.CTkLabel(
            self.info_card,
            text="請購單號: --",
            font=("Microsoft JhengHei UI", 12),
            text_color=dim,
            anchor="w"
        )
        self.lbl_pr_no.pack(fill="x", padx=8, pady=1)

        # 專案名稱
        self.lbl_project = ctk.CTkLabel(
            self.info_card,
            text="專案: --",
            font=("Microsoft JhengHei UI", 11),
            text_color=dim,
            anchor="w",
            justify="left",
            wraplength=170
        )
        self.lbl_project.pack(fill="x", padx=8, pady=(1, 5))

    def create_main_functions_section(self):
        """創建主要功能區"""
        main_label = ctk.CTkLabel(
            self.sidebar,
            text="⚙️ 主要功能",
            font=("Microsoft JhengHei UI", 13, "bold"),
            text_color="#E0E0E0"
        )
        main_label.pack(pady=(0, 4))

        self.btn_get_params = ctk.CTkButton(
            self.sidebar,
            text="📋 從 Odoo 獲取參數",
            command=self.get_parameters_from_odoo,
            height=34,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#388E3C",
            hover_color="#1B5E20",
            text_color="white"
        )
        self.btn_get_params.pack(fill="x", padx=15, pady=3)

        self.btn_push_boq = ctk.CTkButton(
            self.sidebar,
            text="📊 推送到 BOQ",
            command=self.push_to_boq,
            height=34,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#F57C00",
            hover_color="#E65100",
            text_color="white"
        )
        self.btn_push_boq.pack(fill="x", padx=15, pady=3)

        self.btn_transfer_pr = ctk.CTkButton(
            self.sidebar,
            text="🔄 轉移 BOQ 到 PR",
            command=self.transfer_boq_to_pr,
            height=34,
            font=("Microsoft JhengHei UI", 13, "bold"),
            corner_radius=8,
            fg_color="#5D4037",
            hover_color="#3E2723",
            text_color="white"
        )
        self.btn_transfer_pr.pack(fill="x", padx=15, pady=3)
    
    def create_tools_section(self):
        """創建工具功能區"""
        tools_label = ctk.CTkLabel(
            self.sidebar,
            text="🔧 工具",
            font=("Microsoft JhengHei UI", 13, "bold"),
            text_color="#E0E0E0"
        )
        tools_label.pack(pady=(0, 4))

        self.btn_clear_table = ctk.CTkButton(
            self.sidebar,
            text="🗑️ 清除此配置表格ID",
            command=self.clear_table_id,
            height=34,
            font=("Microsoft JhengHei UI", 12, "bold"),
            corner_radius=8,
            fg_color="#FF8F00",
            hover_color="#E65100",
            text_color="white"
        )
        self.btn_clear_table.pack(fill="x", padx=15, pady=3)
        
        self.btn_clear_all_tables = ctk.CTkButton(
            self.sidebar,
            text="🗑️ 清除所有配置表格ID",
            command=self.clear_all_tables_id,
            height=34,
            font=("Microsoft JhengHei UI", 12, "bold"),
            corner_radius=8,
            fg_color="#D32F2F",
            hover_color="#B71C1C",
            text_color="white"
        )
        self.btn_clear_all_tables.pack(fill="x", padx=15, pady=3)

        self._sidebar_sep()

        self.btn_sse_control = ctk.CTkButton(
            self.sidebar,
            text="🤖 MCP 伺服器控制",
            command=self.create_sse_control_panel,
            height=34,
            font=("Microsoft JhengHei UI", 12, "bold"),
            corner_radius=8,
            fg_color="#FF9800",
            hover_color="#F57C00",
            text_color="white"
        )
        self.btn_sse_control.pack(fill="x", padx=15, pady=3)

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
        mode_tag = f" ({self.autocad_mode.upper()})"
        if hasattr(self, 'autocad_util') and self.autocad_util.connected_autocad():
            self.autocad_status_label.configure(
                text=f"📐 AutoCAD{mode_tag}: ✅ 已連接"
            )
            self.btn_connect_autocad.configure(
                fg_color="#2E7D32",
                hover_color="#1B5E20"
            )
        else:
            self.autocad_status_label.configure(
                text=f"📐 AutoCAD{mode_tag}: ❌ 未連接"
            )
            self.btn_connect_autocad.configure(
                fg_color="#7B1FA2",
                hover_color="#4A148C"
            )

        # 更新圖面資訊
        self._update_drawing_info()
    
    def _update_drawing_info(self):
        """更新圖面資訊（配置/請購單號/專案）"""
        if not hasattr(self, 'lbl_layout'):
            return

        connected = hasattr(self, 'autocad_util') and self.autocad_util.connected_autocad()
        active = "#FFFFFF"     # 白色 — 有值時
        accent = "#FFD54F"     # 亮黃色 — PR/專案重要資訊
        dim = "#90A4AE"        # 灰色 — 未連接/無值

        if connected:
            # Layout name — prefer layout_name from block attrs (IPC), fall back to get_active_layout
            layout = getattr(self.autocad_util, 'layout_name', None)
            if not layout:
                try:
                    layout = self.autocad_util.get_active_layout()
                except Exception:
                    pass
            # COM mode may return COM object — extract .Name or convert to str
            if layout and not isinstance(layout, str):
                try:
                    layout = layout.Name if hasattr(layout, 'Name') else str(layout)
                except Exception:
                    layout = None
            if layout:
                self.lbl_layout.configure(text=f"配置 (Layout): {layout}", text_color=active)
            else:
                self.lbl_layout.configure(text="配置 (Layout): --", text_color=dim)

            # PR No
            pr_no = getattr(self.autocad_util, 'pr_no', None)
            if pr_no:
                self.lbl_pr_no.configure(text=f"請購單號: {pr_no}", text_color=accent)
            else:
                self.lbl_pr_no.configure(text="請購單號: --", text_color=dim)

            # Project name
            project_name = getattr(self.autocad_util, 'project_name', None)
            if project_name:
                self.lbl_project.configure(text=f"專案: {project_name}", text_color=active)
            else:
                self.lbl_project.configure(text="專案: --", text_color=dim)
        else:
            self.lbl_layout.configure(text="配置 (Layout): --", text_color=dim)
            self.lbl_pr_no.configure(text="請購單號: --", text_color=dim)
            self.lbl_project.configure(text="專案: --", text_color=dim)

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
            
            # 更新MCP伺服器的AutoCAD dispatcher
            if hasattr(self, 'mcp_manager') and self.mcp_manager:
                try:
                    self.mcp_manager.update_autocad_dispatcher(self.autocad_dispatcher)
                except Exception as e:
                    if self.log_util:
                        self.log_util.safe_log_insert(f"MCP dispatcher 更新失敗: {e}\n")
    
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
        """推送到BOQ — 帶進度條的背景執行"""
        self._run_with_progress(
            title="推送到 BOQ",
            message="準備中...",
            worker=lambda cb: self.push_to_boq_util.push_to_boq(progress_callback=cb),
            success_msg="推送到 BOQ 完成",
            fail_msg="推送到 BOQ 失敗或無資料",
        )

    def transfer_boq_to_pr(self):
        """轉移BOQ到PR — 帶進度條的背景執行"""
        self._run_with_progress(
            title="轉移 BOQ 到 PR",
            message="準備中...",
            worker=lambda cb: self.transfer_boq_to_pr_util.transfer_boq_to_pr(progress_callback=cb),
            success_msg="轉移 BOQ 到 PR 完成",
            fail_msg="轉移 BOQ 到 PR 失敗或無資料",
        )

    def _run_with_progress(self, title, message, worker, success_msg, fail_msg):
        """通用：ProgressDialog + 背景 thread 執行帶進度回調的工作

        Args:
            title: 對話框標題
            message: 初始訊息
            worker: callable(progress_callback) -> bool
            success_msg / fail_msg: 結束時的訊息
        """
        from ui.enhanced_widgets import ProgressDialog

        self.log_util.safe_log_insert(f"[Progress] 開始: {title}\n")
        dialog = ProgressDialog(self, title=title, message=message)
        progress_queue = queue.Queue()

        # Progress callback — 由 worker thread 呼叫，寫入 queue
        def on_progress(value, msg):
            progress_queue.put((value, msg))
        # Attach cancel flag so worker can check it
        on_progress._cancelled = False

        def _worker_thread():
            try:
                self.log_util.safe_log_insert(f"[Progress] Worker thread 啟動\n")
                result = worker(on_progress)
                self.log_util.safe_log_insert(f"[Progress] Worker thread 結束, result={result}\n")
                progress_queue.put(("done", result))
            except Exception as e:
                self.log_util.safe_log_insert(f"[Progress] Worker thread 例外: {e}\n")
                import traceback
                self.log_util.safe_log_insert(f"[Progress] {traceback.format_exc()}\n")
                progress_queue.put(("error", str(e)))

        t = threading.Thread(target=_worker_thread, daemon=True)
        t.start()

        def _poll_progress():
            # Check if dialog was cancelled
            if dialog.cancelled:
                on_progress._cancelled = True
                self.log_util.safe_log_insert(f"[Progress] 使用者取消: {title}\n")
                return

            try:
                while True:
                    item = progress_queue.get_nowait()
                    if item[0] == "done":
                        dialog.update_progress(1.0, "完成")
                        dialog.grab_release()
                        dialog.destroy()
                        if item[1]:
                            self.log_util.safe_log_insert(f"[Progress] ✔ 成功: {success_msg}\n")
                            self.show_info_message("完成", success_msg)
                        else:
                            self.log_util.safe_log_insert(f"[Progress] ⚠ 結束但無資料: {fail_msg}\n")
                            self.show_info_message("提示", fail_msg)
                        return
                    elif item[0] == "error":
                        dialog.grab_release()
                        dialog.destroy()
                        self.log_util.safe_log_insert(f"[Progress] ✘ 錯誤: {item[1]}\n")
                        self.show_error_message("錯誤", f"執行失敗: {item[1]}")
                        return
                    else:
                        value, msg = item
                        dialog.update_progress(value, msg)
            except queue.Empty:
                pass
            # Keep polling every 100ms
            self.after(100, _poll_progress)

        self.after(100, _poll_progress)
    
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
        """初始化MCP服務管理器（向後相容）"""
        try:
            self.log_util.safe_log_insert("MCP服務管理器已準備就緒\n")
            self.log_util.safe_log_insert(f"MCP伺服器端口: {self.mcp_manager.port}\n")
        except Exception as e:
            self.log_util.safe_log_insert(f"MCP服務管理器初始化失敗: {e}\n")
            messagebox.showerror("錯誤", f"無法初始化AI助手服務管理器：{e}")
    
    # toggle_mcp_server 方法已移除 - MCP 按鈕已從 UI 中移除
    
    # update_mcp_status_display 方法已移除 - MCP 按鈕已從 UI 中移除
    
    def toggle_sse_server(self):
        """切換 MCP 伺服器狀態"""
        try:
            if self.mcp_manager.is_running:
                self.mcp_manager.stop_server()
            else:
                self.mcp_manager.start_server()
        except Exception as e:
            self.log_util.safe_log_insert(f"[MCP] Toggle server error: {e}\n")
            messagebox.showerror("錯誤", f"無法切換 MCP 伺服器狀態：{e}")

    def auto_start_mcp_server(self):
        """自動啟動 MCP 伺服器"""
        try:
            self.log_util.safe_log_insert(f"[MCP] Auto-starting MCP server on port {self.mcp_manager.port}\n")
            self.mcp_manager.start_server()
        except Exception as e:
            self.log_util.safe_log_insert(f"[MCP] Auto-start failed: {e}\n")

    # Keep old name for backward compatibility
    auto_start_sse_server = auto_start_mcp_server

    def on_mcp_status_update(self, is_running: bool, message: str):
        """MCP 狀態更新回調"""
        def update_ui():
            if is_running:
                if hasattr(self, 'sse_status_label'):
                    self.sse_status_label.configure(text="🟢")
                if hasattr(self, 'sse_toggle_button'):
                    self.sse_toggle_button.configure(text="⏹️")
                if hasattr(self, 'sse_info_label'):
                    self.sse_info_label.configure(text=f"MCP: :{self.mcp_manager.port}")
            else:
                if hasattr(self, 'sse_status_label'):
                    self.sse_status_label.configure(text="🔴")
                if hasattr(self, 'sse_toggle_button'):
                    self.sse_toggle_button.configure(text="🌊")
                if hasattr(self, 'sse_info_label'):
                    self.sse_info_label.configure(text="")
            self.log_util.safe_log_insert(f"[MCP] {message}\n")
            if is_running:
                self._log_mcp_connection_guide()
        self.after(0, update_ui)

    def _log_mcp_connection_guide(self):
        """在 log 顯示 MCP 連線設定指引"""
        port = self.mcp_manager.port
        endpoint = "/mcp" if self.mcp_manager.transport == "streamable-http" else "/sse"
        url = f"http://localhost:{port}{endpoint}"
        # Python executable path for stdio mode
        python_exe = sys.executable.replace("\\", "/")
        server_script = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "mcp_server_autocad.py"
        ).replace("\\", "/")
        log = self.log_util.safe_log_insert
        log("─" * 60 + "\n")
        log(f"📋 MCP 連線設定 (共 13 個工具, port {port})\n")
        log("─" * 60 + "\n")
        log("方式一：Streamable HTTP（推薦，GUI 已自動啟動）\n")
        log(f"  URL: {url}\n\n")
        log("【Claude Code .mcp.json】\n")
        log('  {\n')
        log('    "mcpServers": {\n')
        log('      "autocad-odoo": {\n')
        log(f'        "type": "http",\n'
            f'      "url": "{url}"\n')
        log('      }\n')
        log('    }\n')
        log('  }\n\n')
        log("【Gemini CLI ~/.gemini/settings.json】\n")
        log('  {\n')
        log('    "mcpServers": {\n')
        log('      "autocad-odoo": {\n')
        log(f'        "type": "http",\n'
            f'      "url": "{url}"\n')
        log('      }\n')
        log('    }\n')
        log('  }\n\n')
        log("─" * 40 + "\n")
        log("方式二：stdio（不需 GUI，獨立執行）\n\n")
        log("【Claude Code .mcp.json】\n")
        log('  {\n')
        log('    "mcpServers": {\n')
        log('      "autocad-odoo": {\n')
        log(f'        "command": "{python_exe}",\n')
        log(f'        "args": ["{server_script}"]\n')
        log('      }\n')
        log('    }\n')
        log('  }\n\n')
        log("【Gemini CLI ~/.gemini/settings.json】\n")
        log('  {\n')
        log('    "mcpServers": {\n')
        log('      "autocad-odoo": {\n')
        log(f'        "command": "{python_exe}",\n')
        log(f'        "args": ["{server_script}"]\n')
        log('      }\n')
        log('    }\n')
        log('  }\n')
        log("─" * 60 + "\n")

    # Keep old callback name for backward compat
    def on_mcp_sse_status_update(self, message: str, is_running: bool):
        """SSE 狀態更新回調 (backward compat)"""
        self.on_mcp_status_update(is_running, message)
    
    def show_sse_status(self):
        """顯示 MCP 伺服器詳細狀態"""
        status_text = f"""
MCP 伺服器狀態:
運行狀態: {'🟢 運行中' if self.mcp_manager.is_running else '🔴 已停止'}
端口: {self.mcp_manager.port}
AutoCAD 模式: {self.autocad_mode.upper()}
        """
        messagebox.showinfo("MCP 伺服器狀態", status_text)

    def test_sse_connection(self):
        """測試 MCP 連接"""
        status = "Running" if self.mcp_manager.is_running else "Stopped"
        messagebox.showinfo("MCP 連接測試", f"MCP Server: {status}\nPort: {self.mcp_manager.port}")
    
    def create_sse_control_panel(self):
        """創建 MCP 控制面板"""
        # 清除主要內容（移除歡迎畫面）
        for widget in self.main_content.winfo_children():
            widget.destroy()

        panel_frame = ctk.CTkFrame(self.main_content, fg_color="transparent")
        panel_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 頂部：標題 + 狀態 + 按鈕（緊湊排列）
        top_row = ctk.CTkFrame(panel_frame, fg_color="transparent")
        top_row.pack(fill="x", pady=(5, 5))

        ctk.CTkLabel(
            top_row, text="🤖 MCP 伺服器",
            font=("Microsoft JhengHei UI", 16, "bold"),
            text_color=theme.get_color('text_primary')
        ).pack(side="left", padx=(5, 10))

        self.sse_panel_status_label = ctk.CTkLabel(
            top_row, text="🟢 運行中",
            font=("Microsoft JhengHei UI", 13),
            text_color="#4CAF50"
        )
        self.sse_panel_status_label.pack(side="left", padx=5)

        # 按鈕靠右
        self.sse_panel_toggle_button = ctk.CTkButton(
            top_row, text="⏹️ 停止", command=self.toggle_sse_server,
            height=30, width=80, font=("Microsoft JhengHei UI", 12),
            corner_radius=6, fg_color="#f44336", hover_color="#d32f2f"
        )
        self.sse_panel_toggle_button.pack(side="right", padx=3)

        ctk.CTkButton(
            top_row, text="📊 狀態", command=self.show_sse_status,
            height=30, width=80, font=("Microsoft JhengHei UI", 12),
            corner_radius=6, fg_color="#9C27B0", hover_color="#7B1FA2"
        ).pack(side="right", padx=3)

        ctk.CTkButton(
            top_row, text="🧪 測試", command=self.test_sse_connection,
            height=30, width=80, font=("Microsoft JhengHei UI", 12),
            corner_radius=6, fg_color="#2196F3", hover_color="#1976D2"
        ).pack(side="right", padx=3)

        # 配置文字框（佔滿剩餘空間）
        port = self.mcp_manager.port
        transport = self.mcp_manager.transport
        endpoint = "/mcp" if transport == "streamable-http" else "/sse"
        url = f"http://localhost:{port}{endpoint}"
        python_exe = sys.executable.replace("\\", "/")
        server_script = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "mcp_server_autocad.py"
        ).replace("\\", "/")

        config_text = (
            f"MCP Server: {url}  |  Transport: {transport}  |  "
            f"AutoCAD: {self.autocad_mode.upper()}  |  工具數: 13\n"
            f"{'─' * 70}\n"
            f"\n"
            f"方式一：Streamable HTTP（推薦，GUI 已自動啟動）\n"
            f"\n"
            f"【Claude Code】專案目錄建立 .mcp.json:\n"
            f'{{\n'
            f'  "mcpServers": {{\n'
            f'    "autocad-odoo": {{\n'
            f'      "type": "http",\n'
            f'      "url": "{url}"\n'
            f'    }}\n'
            f'  }}\n'
            f'}}\n'
            f"\n"
            f"【Gemini CLI】~/.gemini/settings.json:\n"
            f'{{\n'
            f'  "mcpServers": {{\n'
            f'    "autocad-odoo": {{\n'
            f'      "type": "http",\n'
            f'      "url": "{url}"\n'
            f'    }}\n'
            f'  }}\n'
            f'}}\n'
            f"\n"
            f"{'─' * 70}\n"
            f"\n"
            f"方式二：stdio（不需 GUI，AI 工具自動啟動）\n"
            f"\n"
            f"【Claude Code / Gemini CLI】.mcp.json 或 settings.json:\n"
            f'{{\n'
            f'  "mcpServers": {{\n'
            f'    "autocad-odoo": {{\n'
            f'      "command": "{python_exe}",\n'
            f'      "args": ["{server_script}"]\n'
            f'    }}\n'
            f'  }}\n'
            f'}}\n'
        )

        config_box = ctk.CTkTextbox(
            panel_frame,
            font=("Consolas", 13),
            wrap="none"
        )
        config_box.pack(fill="both", expand=True, pady=(5, 5))
        config_box.insert("0.0", config_text)
        config_box.configure(state="disabled")

        # 更新面板狀態
        self.update_sse_panel_status()
    
    def update_sse_panel_status(self):
        """更新 MCP 面板狀態"""
        if hasattr(self, 'sse_panel_status_label'):
            is_running = self.mcp_manager.is_running
            if is_running:
                self.sse_panel_status_label.configure(text="狀態: 🟢 運行中")
                if hasattr(self, 'sse_panel_toggle_button'):
                    self.sse_panel_toggle_button.configure(
                        text="⏹️ 停止 MCP 伺服器",
                        fg_color="#f44336",  # 紅色停止按鈕
                        hover_color="#d32f2f"
                    )
            else:
                self.sse_panel_status_label.configure(text="狀態: 🔴 已停止")
                if hasattr(self, 'sse_panel_toggle_button'):
                    self.sse_panel_toggle_button.configure(
                        text="🚀 啟動 MCP 伺服器",
                        fg_color="#4CAF50",  # 綠色啟動按鈕
                        hover_color="#388E3C"
                    )
    
    def start_gui_proxy_processing(self):
        """啟動GUI代理請求處理"""
        self.process_gui_proxy_requests()
    
    def process_gui_proxy_requests(self):
        """處理GUI代理請求 (在主線程中運行, COM模式only)"""
        if self.gui_proxy is None:
            return
        try:
            processed = self.gui_proxy.process_requests()
            if processed > 0:
                self.log_util.safe_log_insert(f"[GUI Proxy] 處理了 {processed} 個請求\n")
        except Exception as e:
            self.log_util.safe_log_insert(f"[GUI Proxy] 處理請求時發生錯誤: {e}\n")
        # 每100毫秒檢查一次
        self.after(100, self.process_gui_proxy_requests)
    
    def on_closing(self):
        """視窗關閉事件"""
        # 停止 MCP 伺服器
        try:
            if hasattr(self, 'mcp_manager') and self.mcp_manager and self.mcp_manager.is_running:
                self.mcp_manager.stop_server()
                self.log_util.safe_log_insert("MCP 伺服器已停止\n")
        except Exception as e:
            self.log_util.safe_log_insert(f"停止 MCP 伺服器失敗: {e}\n")

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