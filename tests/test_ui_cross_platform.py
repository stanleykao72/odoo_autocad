# -*- coding: utf-8 -*-
"""
跨平台UI測試 - 避免Windows依賴項
"""
import sys
import os
import customtkinter as ctk
from tkinter import messagebox
from typing import Optional

# 確保可以導入專案模組
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 跨平台UI基礎類
class CrossPlatformUITest(ctk.CTk):
    """跨平台的現代化UI測試"""
    
    def __init__(self):
        super().__init__()
        
        # 設置基本主題
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        self.title("AutoCAD Odoo 整合系統 - UI測試")
        self.geometry("1200x800")
        
        # 設置視窗圖示（如果存在）
        try:
            self.iconpath = "icon/odoo_autocad.ico"
            if os.path.exists(self.iconpath):
                self.iconbitmap(self.iconpath)
        except:
            pass
        
        # 創建UI
        self.create_ui()
    
    def create_ui(self):
        """創建主要UI元件"""
        
        # 主要容器 - 使用grid佈局
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # 左側導航欄
        self.create_sidebar()
        
        # 主要內容區域
        self.create_main_content()
        
        # 底部狀態欄
        self.create_status_bar()
    
    def create_sidebar(self):
        """創建左側導航欄"""
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)
        
        # 標題
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="AutoCAD Odoo",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # 連接狀態指示器
        self.create_connection_indicators()
        
        # 功能按鈕
        self.create_function_buttons()
        
        # 外觀模式選擇器
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="外觀模式:", anchor="w")
        self.appearance_mode_label.grid(row=9, column=0, padx=20, pady=(10, 0))
        
        self.appearance_mode_menu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["Light", "Dark", "System"],
            command=self.change_appearance_mode_event
        )
        self.appearance_mode_menu.grid(row=10, column=0, padx=20, pady=(10, 20))
    
    def create_connection_indicators(self):
        """創建連接狀態指示器"""
        # Odoo連接狀態
        self.odoo_status_frame = ctk.CTkFrame(self.sidebar_frame)
        self.odoo_status_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.odoo_status_label = ctk.CTkLabel(
            self.odoo_status_frame, 
            text="🔗 Odoo狀態"
        )
        self.odoo_status_label.grid(row=0, column=0, padx=10, pady=5)
        
        self.odoo_status_indicator = ctk.CTkLabel(
            self.odoo_status_frame,
            text="● 離線",
            text_color="red"
        )
        self.odoo_status_indicator.grid(row=1, column=0, padx=10, pady=5)
        
        # AutoCAD連接狀態
        self.autocad_status_frame = ctk.CTkFrame(self.sidebar_frame)
        self.autocad_status_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        self.autocad_status_label = ctk.CTkLabel(
            self.autocad_status_frame,
            text="🎯 AutoCAD狀態"
        )
        self.autocad_status_label.grid(row=0, column=0, padx=10, pady=5)
        
        self.autocad_status_indicator = ctk.CTkLabel(
            self.autocad_status_frame,
            text="● 未連接",
            text_color="orange"
        )
        self.autocad_status_indicator.grid(row=1, column=0, padx=10, pady=5)
    
    def create_function_buttons(self):
        """創建功能按鈕"""
        functions = [
            ("📊 同步產品資料", self.sync_products),
            ("📋 參數輸入", self.open_param_form),
            ("📄 推送至BOQ", self.push_to_boq),
            ("🛒 轉換為PR", self.transfer_to_pr),
            ("⚙️ 系統設定", self.open_settings),
        ]
        
        for i, (text, command) in enumerate(functions, start=3):
            button = ctk.CTkButton(
                self.sidebar_frame,
                text=text,
                command=command,
                width=200
            )
            button.grid(row=i, column=0, padx=20, pady=10)
    
    def create_main_content(self):
        """創建主要內容區域"""
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        
        # 標題
        self.main_title = ctk.CTkLabel(
            self.main_frame,
            text="歡迎使用 AutoCAD Odoo 整合系統",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.main_title.grid(row=0, column=0, padx=20, pady=20)
        
        # 內容標籤頁
        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        # 儀表板標籤
        self.tabview.add("儀表板")
        self.create_dashboard_tab()
        
        # 產品資料標籤
        self.tabview.add("產品資料")
        self.create_products_tab()
        
        # 日誌標籤
        self.tabview.add("系統日誌")
        self.create_logs_tab()
    
    def create_dashboard_tab(self):
        """創建儀表板標籤內容"""
        dashboard_frame = self.tabview.tab("儀表板")
        
        # 統計卡片
        stats_frame = ctk.CTkFrame(dashboard_frame)
        stats_frame.pack(fill="x", padx=20, pady=20)
        
        # 統計項目
        stats = [
            ("已同步產品", "156", "green"),
            ("待處理BOQ", "23", "orange"),
            ("已完成PR", "45", "blue")
        ]
        
        for i, (title, value, color) in enumerate(stats):
            card = ctk.CTkFrame(stats_frame)
            card.grid(row=0, column=i, padx=10, pady=10, sticky="ew")
            stats_frame.grid_columnconfigure(i, weight=1)
            
            title_label = ctk.CTkLabel(card, text=title)
            title_label.pack(pady=(10, 5))
            
            value_label = ctk.CTkLabel(
                card, 
                text=value, 
                font=ctk.CTkFont(size=28, weight="bold"),
                text_color=color
            )
            value_label.pack(pady=(0, 10))
        
        # 快速操作區域
        quick_actions_frame = ctk.CTkFrame(dashboard_frame)
        quick_actions_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        quick_title = ctk.CTkLabel(
            quick_actions_frame,
            text="快速操作",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        quick_title.pack(pady=20)
        
        actions_container = ctk.CTkFrame(quick_actions_frame)
        actions_container.pack(fill="x", padx=20, pady=20)
        
        quick_actions = [
            ("🔄 重新連接", self.reconnect_services),
            ("📥 匯入圖檔", self.import_drawing),
            ("📤 匯出報表", self.export_report)
        ]
        
        for i, (text, command) in enumerate(quick_actions):
            button = ctk.CTkButton(
                actions_container,
                text=text,
                command=command,
                width=150,
                height=50
            )
            button.grid(row=0, column=i, padx=10, pady=10)
            actions_container.grid_columnconfigure(i, weight=1)
    
    def create_products_tab(self):
        """創建產品資料標籤內容"""
        products_frame = self.tabview.tab("產品資料")
        
        # 搜尋框
        search_frame = ctk.CTkFrame(products_frame)
        search_frame.pack(fill="x", padx=20, pady=20)
        
        search_label = ctk.CTkLabel(search_frame, text="搜尋產品:")
        search_label.pack(side="left", padx=10, pady=10)
        
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="輸入產品名稱或編號...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        
        search_button = ctk.CTkButton(search_frame, text="搜尋", command=self.search_products)
        search_button.pack(side="right", padx=10, pady=10)
        
        # 模擬產品列表
        products_list_frame = ctk.CTkScrollableFrame(products_frame)
        products_list_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 表頭
        headers = ["產品編號", "產品名稱", "規格", "單位", "狀態"]
        for i, header in enumerate(headers):
            label = ctk.CTkLabel(
                products_list_frame,
                text=header,
                font=ctk.CTkFont(weight="bold")
            )
            label.grid(row=0, column=i, padx=10, pady=5, sticky="w")
        
        # 模擬資料
        sample_products = [
            ("P001", "螺絲 M6x20", "M6x20mm", "個", "已同步"),
            ("P002", "螺帽 M6", "M6", "個", "已同步"),
            ("P003", "墊片 6mm", "內徑6mm", "個", "待同步"),
        ]
        
        for i, product in enumerate(sample_products, start=1):
            for j, value in enumerate(product):
                label = ctk.CTkLabel(products_list_frame, text=value)
                label.grid(row=i, column=j, padx=10, pady=5, sticky="w")
    
    def create_logs_tab(self):
        """創建系統日誌標籤內容"""
        logs_frame = self.tabview.tab("系統日誌")
        
        # 日誌控制框
        log_controls = ctk.CTkFrame(logs_frame)
        log_controls.pack(fill="x", padx=20, pady=20)
        
        clear_button = ctk.CTkButton(log_controls, text="清除日誌", command=self.clear_logs)
        clear_button.pack(side="left", padx=10, pady=10)
        
        export_logs_button = ctk.CTkButton(log_controls, text="匯出日誌", command=self.export_logs)
        export_logs_button.pack(side="left", padx=10, pady=10)
        
        # 日誌顯示區域
        self.log_textbox = ctk.CTkTextbox(logs_frame)
        self.log_textbox.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 添加示例日誌
        sample_logs = [
            "[2025-06-23 10:30:15] INFO: 應用程式啟動",
            "[2025-06-23 10:30:16] INFO: 載入設定檔案",
            "[2025-06-23 10:30:17] WARNING: Odoo連接失敗，嘗試重新連接...",
            "[2025-06-23 10:30:20] INFO: AutoCAD狀態檢查完成",
            "[2025-06-23 10:30:25] INFO: UI初始化完成"
        ]
        
        for log in sample_logs:
            self.log_textbox.insert("end", log + "\n")
    
    def create_status_bar(self):
        """創建底部狀態欄"""
        self.status_frame = ctk.CTkFrame(self, height=30)
        self.status_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 20))
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="準備就緒 | UI測試模式 | CustomTkinter 5.2.2",
            anchor="w"
        )
        self.status_label.pack(side="left", padx=20, pady=5)
        
        self.version_label = ctk.CTkLabel(
            self.status_frame,
            text="v1.0.0-modern",
            anchor="e"
        )
        self.version_label.pack(side="right", padx=20, pady=5)
    
    # 事件處理函數
    def change_appearance_mode_event(self, new_appearance_mode: str):
        """更改外觀模式"""
        ctk.set_appearance_mode(new_appearance_mode)
        self.log_message(f"外觀模式已切換至: {new_appearance_mode}")
    
    def sync_products(self):
        """同步產品資料"""
        messagebox.showinfo("功能測試", "同步產品資料功能（測試模式）")
        self.log_message("執行產品資料同步")
    
    def open_param_form(self):
        """開啟參數輸入表單"""
        messagebox.showinfo("功能測試", "參數輸入表單（測試模式）")
        self.log_message("開啟參數輸入表單")
    
    def push_to_boq(self):
        """推送至BOQ"""
        messagebox.showinfo("功能測試", "推送至BOQ功能（測試模式）")
        self.log_message("執行BOQ推送")
    
    def transfer_to_pr(self):
        """轉換為PR"""
        messagebox.showinfo("功能測試", "轉換為PR功能（測試模式）")
        self.log_message("執行BOQ轉PR")
    
    def open_settings(self):
        """開啟系統設定"""
        messagebox.showinfo("功能測試", "系統設定（測試模式）")
        self.log_message("開啟系統設定")
    
    def reconnect_services(self):
        """重新連接服務"""
        messagebox.showinfo("功能測試", "重新連接服務（測試模式）")
        self.log_message("重新連接Odoo和AutoCAD服務")
    
    def import_drawing(self):
        """匯入圖檔"""
        messagebox.showinfo("功能測試", "匯入圖檔功能（測試模式）")
        self.log_message("執行圖檔匯入")
    
    def export_report(self):
        """匯出報表"""
        messagebox.showinfo("功能測試", "匯出報表功能（測試模式）")
        self.log_message("匯出系統報表")
    
    def search_products(self):
        """搜尋產品"""
        search_term = self.search_entry.get()
        self.log_message(f"搜尋產品: {search_term}")
        messagebox.showinfo("搜尋結果", f"搜尋「{search_term}」的結果（測試模式）")
    
    def clear_logs(self):
        """清除日誌"""
        self.log_textbox.delete("1.0", "end")
        self.log_message("日誌已清除")
    
    def export_logs(self):
        """匯出日誌"""
        messagebox.showinfo("匯出日誌", "日誌匯出功能（測試模式）")
        self.log_message("執行日誌匯出")
    
    def log_message(self, message: str):
        """添加日誌訊息"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] INFO: {message}\n"
        self.log_textbox.insert("end", log_entry)
        self.log_textbox.see("end")

def main():
    """主函數"""
    print("=== AutoCAD Odoo 整合系統 - 跨平台UI測試 ===")
    print("✅ CustomTkinter UI 測試模式啟動")
    
    try:
        app = CrossPlatformUITest()
        print("✅ 現代化UI創建成功")
        app.mainloop()
    except Exception as e:
        print(f"❌ UI啟動失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()