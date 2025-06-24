# -*- coding: utf-8 -*-
"""
測試增強UI組件
展示所有新的現代化UI功能
"""
import sys
import os
import customtkinter as ctk
from tkinter import messagebox
import threading
import time
import random

# 確保可以導入專案模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.enhanced_widgets import (
    StatusIndicator, StatCard, ActionButton, SearchEntry, 
    LogViewer, DataTable, ProgressDialog
)


class EnhancedUIDemo(ctk.CTk):
    """增強UI功能展示"""
    
    def __init__(self):
        super().__init__()
        
        # 設置基本主題
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        self.title("AutoCAD Odoo 整合系統 - 增強UI展示")
        self.geometry("1400x900")
        
        # 設置網格權重
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # 創建UI
        self.create_ui()
        
        # 模擬數據更新
        self.start_demo_updates()
    
    def create_ui(self):
        """創建主要UI"""
        
        # 左側控制面板
        self.create_control_panel()
        
        # 主要內容區域
        self.create_main_content()
        
        # 底部狀態欄
        self.create_status_bar()
    
    def create_control_panel(self):
        """創建左側控制面板"""
        self.control_frame = ctk.CTkFrame(self, width=300)
        self.control_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.control_frame.grid_propagate(False)
        
        # 標題
        title_label = ctk.CTkLabel(
            self.control_frame,
            text="控制面板",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(pady=20)
        
        # 連接狀態區域
        status_frame = ctk.CTkFrame(self.control_frame)
        status_frame.pack(fill="x", padx=20, pady=10)
        
        status_title = ctk.CTkLabel(
            status_frame,
            text="連接狀態",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        status_title.pack(pady=10)
        
        # Odoo狀態指示器
        self.odoo_status = StatusIndicator(status_frame, "Odoo服務")
        self.odoo_status.pack(fill="x", padx=10, pady=5)
        
        # AutoCAD狀態指示器
        self.autocad_status = StatusIndicator(status_frame, "AutoCAD")
        self.autocad_status.pack(fill="x", padx=10, pady=5)
        
        # 統計卡片區域
        stats_frame = ctk.CTkFrame(self.control_frame)
        stats_frame.pack(fill="x", padx=20, pady=10)
        
        stats_title = ctk.CTkLabel(
            stats_frame,
            text="系統統計",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        stats_title.pack(pady=(10, 5))
        
        # 統計卡片
        self.products_card = StatCard(stats_frame, "產品數量", "0", "blue")
        self.products_card.pack(fill="x", padx=10, pady=5)
        
        self.boq_card = StatCard(stats_frame, "BOQ項目", "0", "green")
        self.boq_card.pack(fill="x", padx=10, pady=5)
        
        self.pr_card = StatCard(stats_frame, "PR數量", "0", "orange")
        self.pr_card.pack(fill="x", padx=10, pady=5)
        
        # 快速操作按鈕
        actions_frame = ctk.CTkFrame(self.control_frame)
        actions_frame.pack(fill="x", padx=20, pady=10)
        
        actions_title = ctk.CTkLabel(
            actions_frame,
            text="快速操作",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        actions_title.pack(pady=(10, 5))
        
        # 操作按鈕
        connect_btn = ActionButton(
            actions_frame,
            "連接服務",
            "🔌",
            self.simulate_connection
        )
        connect_btn.pack(fill="x", padx=10, pady=5)
        
        sync_btn = ActionButton(
            actions_frame,
            "同步數據",
            "🔄",
            self.simulate_sync
        )
        sync_btn.pack(fill="x", padx=10, pady=5)
        
        process_btn = ActionButton(
            actions_frame,
            "處理BOQ",
            "📊",
            self.simulate_processing
        )
        process_btn.pack(fill="x", padx=10, pady=5)
        
        # 測試按鈕
        test_frame = ctk.CTkFrame(self.control_frame)
        test_frame.pack(fill="x", padx=20, pady=10)
        
        test_title = ctk.CTkLabel(
            test_frame,
            text="測試功能",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        test_title.pack(pady=(10, 5))
        
        progress_btn = ctk.CTkButton(
            test_frame,
            text="📈 顯示進度",
            command=self.show_progress_dialog
        )
        progress_btn.pack(fill="x", padx=10, pady=5)
        
        log_btn = ctk.CTkButton(
            test_frame,
            text="📝 添加日誌",
            command=self.add_sample_log
        )
        log_btn.pack(fill="x", padx=10, pady=5)
    
    def create_main_content(self):
        """創建主要內容區域"""
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        
        # 標題和搜尋區域
        header_frame = ctk.CTkFrame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=20)
        header_frame.grid_columnconfigure(1, weight=1)
        
        main_title = ctk.CTkLabel(
            header_frame,
            text="數據管理中心",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        main_title.grid(row=0, column=0, sticky="w", padx=20, pady=20)
        
        # 搜尋框
        self.search_entry = SearchEntry(
            header_frame,
            placeholder="搜尋產品、BOQ或PR...",
            search_callback=self.on_search
        )
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=20, pady=20)
        
        # 主要標籤頁
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        
        # 產品數據標籤
        self.create_products_tab()
        
        # 日誌標籤
        self.create_logs_tab()
        
        # 設置標籤
        self.create_settings_tab()
    
    def create_products_tab(self):
        """創建產品數據標籤"""
        products_tab = self.tabview.add("產品數據")
        
        # 產品表格
        headers = ["產品編號", "產品名稱", "規格", "單位", "庫存", "狀態"]
        self.products_table = DataTable(products_tab, headers)
        self.products_table.pack(fill="both", expand=True)
        
        # 添加示例數據
        sample_products = [
            ["P001", "螺絲 M6x20", "M6x20mm", "個", "1000", "正常"],
            ["P002", "螺帽 M6", "M6", "個", "500", "正常"],
            ["P003", "墊片 6mm", "內徑6mm", "個", "200", "庫存不足"],
            ["P004", "角鋼 50x50", "50x50x5mm", "支", "150", "正常"],
            ["P005", "鋼板 10mm", "10mm厚", "片", "50", "正常"],
        ]
        
        for product in sample_products:
            self.products_table.add_row(product)
    
    def create_logs_tab(self):
        """創建日誌標籤"""
        logs_tab = self.tabview.add("系統日誌")
        
        # 日誌查看器
        self.log_viewer = LogViewer(logs_tab)
        self.log_viewer.pack(fill="both", expand=True)
        
        # 添加初始日誌
        self.log_viewer.add_log("INFO", "應用程式啟動")
        self.log_viewer.add_log("INFO", "UI組件初始化完成")
        self.log_viewer.add_log("WARNING", "Odoo連接尚未建立")
    
    def create_settings_tab(self):
        """創建設置標籤"""
        settings_tab = self.tabview.add("系統設置")
        
        # 外觀設置
        appearance_frame = ctk.CTkFrame(settings_tab)
        appearance_frame.pack(fill="x", padx=20, pady=20)
        
        appearance_title = ctk.CTkLabel(
            appearance_frame,
            text="外觀設置",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        appearance_title.pack(pady=10)
        
        # 主題選擇
        theme_frame = ctk.CTkFrame(appearance_frame)
        theme_frame.pack(fill="x", padx=20, pady=10)
        
        theme_label = ctk.CTkLabel(theme_frame, text="主題模式:")
        theme_label.pack(side="left", padx=10, pady=10)
        
        self.theme_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=["Light", "Dark", "System"],
            command=self.change_theme
        )
        self.theme_menu.pack(side="left", padx=10, pady=10)
        
        # 顏色主題選擇
        color_frame = ctk.CTkFrame(appearance_frame)
        color_frame.pack(fill="x", padx=20, pady=10)
        
        color_label = ctk.CTkLabel(color_frame, text="顏色主題:")
        color_label.pack(side="left", padx=10, pady=10)
        
        self.color_menu = ctk.CTkOptionMenu(
            color_frame,
            values=["blue", "green", "dark-blue"],
            command=self.change_color_theme
        )
        self.color_menu.pack(side="left", padx=10, pady=10)
        
        # 連接設置
        connection_frame = ctk.CTkFrame(settings_tab)
        connection_frame.pack(fill="x", padx=20, pady=20)
        
        connection_title = ctk.CTkLabel(
            connection_frame,
            text="連接設置",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        connection_title.pack(pady=10)
        
        # Odoo設置
        odoo_frame = ctk.CTkFrame(connection_frame)
        odoo_frame.pack(fill="x", padx=20, pady=10)
        
        odoo_label = ctk.CTkLabel(odoo_frame, text="Odoo伺服器:")
        odoo_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        self.odoo_entry = ctk.CTkEntry(odoo_frame, placeholder_text="http://localhost:8069")
        self.odoo_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        odoo_frame.grid_columnconfigure(1, weight=1)
        
        # 測試連接按鈕
        test_conn_btn = ctk.CTkButton(
            connection_frame,
            text="測試連接",
            command=self.test_connections
        )
        test_conn_btn.pack(pady=10)
    
    def create_status_bar(self):
        """創建底部狀態欄"""
        self.status_frame = ctk.CTkFrame(self, height=40)
        self.status_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="就緒 | 增強UI演示模式",
            anchor="w"
        )
        self.status_label.pack(side="left", padx=20, pady=10)
        
        self.version_label = ctk.CTkLabel(
            self.status_frame,
            text="v2.0.0-enhanced",
            anchor="e"
        )
        self.version_label.pack(side="right", padx=20, pady=10)
    
    def start_demo_updates(self):
        """開始演示數據更新"""
        def update_loop():
            while True:
                try:
                    # 隨機更新統計數據
                    products_count = random.randint(100, 999)
                    boq_count = random.randint(10, 99)
                    pr_count = random.randint(5, 50)
                    
                    self.after(0, lambda: self.products_card.update_value(str(products_count)))
                    self.after(0, lambda: self.boq_card.update_value(str(boq_count)))
                    self.after(0, lambda: self.pr_card.update_value(str(pr_count)))
                    
                    # 隨機更新連接狀態
                    if random.random() > 0.7:
                        status = random.choice(["online", "offline", "connecting"])
                        self.after(0, lambda s=status: self.odoo_status.set_status(s))
                    
                    if random.random() > 0.8:
                        status = random.choice(["online", "offline", "error"])
                        self.after(0, lambda s=status: self.autocad_status.set_status(s))
                    
                    time.sleep(3)
                except:
                    break
        
        thread = threading.Thread(target=update_loop, daemon=True)
        thread.start()
    
    # 事件處理函數
    def on_search(self, query: str):
        """搜尋處理"""
        self.log_viewer.add_log("INFO", f"搜尋: {query}")
        self.status_label.configure(text=f"搜尋: {query}" if query else "就緒 | 增強UI演示模式")
    
    def simulate_connection(self):
        """模擬連接過程"""
        def connect_process():
            self.odoo_status.set_status("connecting", "正在連接 Odoo...")
            self.autocad_status.set_status("connecting", "正在連接 AutoCAD...")
            time.sleep(2)
            
            self.after(0, lambda: self.odoo_status.set_status("online", "Odoo 連接成功"))
            self.after(0, lambda: self.autocad_status.set_status("online", "AutoCAD 連接成功"))
            self.after(0, lambda: self.log_viewer.add_log("INFO", "所有服務連接成功"))
        
        thread = threading.Thread(target=connect_process, daemon=True)
        thread.start()
    
    def simulate_sync(self):
        """模擬數據同步"""
        def sync_process():
            self.after(0, lambda: self.log_viewer.add_log("INFO", "開始數據同步"))
            time.sleep(1)
            self.after(0, lambda: self.log_viewer.add_log("INFO", "同步產品數據..."))
            time.sleep(1)
            self.after(0, lambda: self.log_viewer.add_log("INFO", "同步BOQ數據..."))
            time.sleep(1)
            self.after(0, lambda: self.log_viewer.add_log("INFO", "數據同步完成"))
        
        thread = threading.Thread(target=sync_process, daemon=True)
        thread.start()
    
    def simulate_processing(self):
        """模擬BOQ處理"""
        self.log_viewer.add_log("INFO", "開始處理BOQ數據")
        self.log_viewer.add_log("WARNING", "發現3個需要確認的項目")
        self.log_viewer.add_log("INFO", "BOQ處理完成")
    
    def show_progress_dialog(self):
        """顯示進度對話框"""
        progress_dialog = ProgressDialog(self, "數據處理", "正在處理數據，請稍候...")
        
        def progress_task():
            for i in range(101):
                if progress_dialog.is_cancelled():
                    break
                
                progress = i / 100
                message = f"處理中... {i}%"
                self.after(0, lambda p=progress, m=message: progress_dialog.update_progress(p, m))
                time.sleep(0.05)
            
            if not progress_dialog.is_cancelled():
                self.after(0, progress_dialog.destroy)
                self.after(0, lambda: self.log_viewer.add_log("INFO", "數據處理完成"))
        
        thread = threading.Thread(target=progress_task, daemon=True)
        thread.start()
    
    def add_sample_log(self):
        """添加示例日誌"""
        import random
        
        levels = ["INFO", "WARNING", "ERROR"]
        messages = [
            "用戶操作記錄",
            "系統狀態檢查",
            "數據庫連接測試",
            "文件處理完成",
            "網絡連接異常",
            "數據驗證失敗",
            "用戶登入成功"
        ]
        
        level = random.choice(levels)
        message = random.choice(messages)
        self.log_viewer.add_log(level, message)
    
    def change_theme(self, theme: str):
        """更改主題"""
        ctk.set_appearance_mode(theme)
        self.log_viewer.add_log("INFO", f"主題已切換至: {theme}")
    
    def change_color_theme(self, color: str):
        """更改顏色主題"""
        ctk.set_default_color_theme(color)
        self.log_viewer.add_log("INFO", f"顏色主題已切換至: {color}")
        messagebox.showinfo("提示", "顏色主題將在重新啟動後生效")
    
    def test_connections(self):
        """測試連接"""
        self.log_viewer.add_log("INFO", "開始測試連接...")
        url = self.odoo_entry.get() or "http://localhost:8069"
        self.log_viewer.add_log("INFO", f"測試Odoo連接: {url}")
        self.log_viewer.add_log("WARNING", "連接測試功能尚未實現")


def main():
    """主函數"""
    print("=== AutoCAD Odoo 整合系統 - 增強UI演示 ===")
    print("✅ 展示現代化UI組件功能")
    
    try:
        app = EnhancedUIDemo()
        print("✅ 增強UI演示啟動成功")
        app.mainloop()
    except Exception as e:
        print(f"❌ UI啟動失敗: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()