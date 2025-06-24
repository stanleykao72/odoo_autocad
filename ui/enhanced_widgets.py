# -*- coding: utf-8 -*-
"""
增強的UI組件
提供現代化的自定義控件
"""
import customtkinter as ctk
from tkinter import ttk
from typing import Callable, Optional, Union, List
import datetime


class StatusIndicator(ctk.CTkFrame):
    """狀態指示器組件"""
    
    def __init__(self, parent, title: str, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.title = title
        self.status = "offline"
        
        # 標題標籤
        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.title_label.pack(pady=(10, 5))
        
        # 狀態指示器
        self.status_frame = ctk.CTkFrame(self)
        self.status_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.status_dot = ctk.CTkLabel(
            self.status_frame,
            text="●",
            font=ctk.CTkFont(size=16),
            text_color="red"
        )
        self.status_dot.pack(side="left", padx=(10, 5), pady=5)
        
        self.status_text = ctk.CTkLabel(
            self.status_frame,
            text="離線",
            font=ctk.CTkFont(size=11)
        )
        self.status_text.pack(side="left", pady=5)
        
        # 最後更新時間
        self.time_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=9),
            text_color="gray"
        )
        self.time_label.pack(side="right", padx=(5, 10), pady=5)
    
    def set_status(self, status: str, message: str = None):
        """設置狀態"""
        self.status = status
        
        status_config = {
            "online": {"color": "green", "text": "線上", "message": message or "連接正常"},
            "offline": {"color": "red", "text": "離線", "message": message or "未連接"},
            "connecting": {"color": "orange", "text": "連接中", "message": message or "正在連接..."},
            "error": {"color": "red", "text": "錯誤", "message": message or "連接錯誤"}
        }
        
        config = status_config.get(status, status_config["offline"])
        
        self.status_dot.configure(text_color=config["color"])
        self.status_text.configure(text=config["message"])
        
        # 更新時間
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.time_label.configure(text=current_time)


class StatCard(ctk.CTkFrame):
    """統計卡片組件"""
    
    def __init__(self, parent, title: str, value: str, color: str = "blue", **kwargs):
        super().__init__(parent, **kwargs)
        
        # 標題
        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.title_label.pack(pady=(15, 5))
        
        # 數值
        self.value_label = ctk.CTkLabel(
            self,
            text=value,
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color=color
        )
        self.value_label.pack(pady=(0, 15))
    
    def update_value(self, value: str, color: str = None):
        """更新數值"""
        self.value_label.configure(text=value)
        if color:
            self.value_label.configure(text_color=color)


class ActionButton(ctk.CTkButton):
    """增強的動作按鈕"""
    
    def __init__(self, parent, text: str, icon: str = None, command: Callable = None, **kwargs):
        # 準備按鈕文字
        button_text = f"{icon} {text}" if icon else text
        
        super().__init__(
            parent,
            text=button_text,
            command=command,
            **kwargs
        )
        
        self.original_command = command
        self.is_loading = False
        self.original_text = button_text
    
    def set_loading(self, loading: bool = True):
        """設置載入狀態"""
        self.is_loading = loading
        if loading:
            self.configure(text="⏳ 處理中...", state="disabled")
        else:
            self.configure(text=self.original_text, state="normal")
    
    def execute_with_loading(self):
        """執行命令並顯示載入狀態"""
        if self.original_command and not self.is_loading:
            self.set_loading(True)
            try:
                self.original_command()
            finally:
                self.after(1000, lambda: self.set_loading(False))


class SearchEntry(ctk.CTkFrame):
    """搜尋輸入框組件"""
    
    def __init__(self, parent, placeholder: str = "搜尋...", search_callback: Callable = None, **kwargs):
        super().__init__(parent, **kwargs)
        
        self.search_callback = search_callback
        
        # 搜尋圖示
        self.search_icon = ctk.CTkLabel(
            self,
            text="🔍",
            font=ctk.CTkFont(size=14)
        )
        self.search_icon.pack(side="left", padx=(10, 5), pady=10)
        
        # 輸入框
        self.entry = ctk.CTkEntry(
            self,
            placeholder_text=placeholder,
            border_width=0
        )
        self.entry.pack(side="left", fill="x", expand=True, pady=10)
        self.entry.bind("<Return>", self._on_search)
        self.entry.bind("<KeyRelease>", self._on_key_release)
        
        # 清除按鈕
        self.clear_button = ctk.CTkButton(
            self,
            text="✕",
            width=30,
            command=self.clear,
            fg_color="transparent",
            text_color="gray",
            hover_color="lightgray"
        )
        self.clear_button.pack(side="right", padx=(5, 10), pady=10)
        self.clear_button.pack_forget()  # 初始隱藏
    
    def _on_search(self, event=None):
        """執行搜尋"""
        if self.search_callback:
            self.search_callback(self.entry.get())
    
    def _on_key_release(self, event=None):
        """按鍵釋放事件"""
        text = self.entry.get()
        if text:
            self.clear_button.pack(side="right", padx=(5, 10), pady=10)
        else:
            self.clear_button.pack_forget()
    
    def clear(self):
        """清除輸入"""
        self.entry.delete(0, "end")
        self.clear_button.pack_forget()
        if self.search_callback:
            self.search_callback("")
    
    def get(self) -> str:
        """獲取輸入值"""
        return self.entry.get()
    
    def set(self, value: str):
        """設置輸入值"""
        self.entry.delete(0, "end")
        self.entry.insert(0, value)


class LogViewer(ctk.CTkFrame):
    """日誌查看器組件"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        
        # 工具欄
        self.toolbar = ctk.CTkFrame(self)
        self.toolbar.pack(fill="x", padx=10, pady=(10, 0))
        
        # 日誌級別過濾器
        self.level_var = ctk.StringVar(value="ALL")
        self.level_menu = ctk.CTkOptionMenu(
            self.toolbar,
            values=["ALL", "INFO", "WARNING", "ERROR"],
            variable=self.level_var,
            command=self._filter_logs,
            width=100
        )
        self.level_menu.pack(side="left", padx=(10, 5), pady=5)
        
        # 自動滾動開關
        self.auto_scroll_var = ctk.BooleanVar(value=True)
        self.auto_scroll_check = ctk.CTkCheckBox(
            self.toolbar,
            text="自動滾動",
            variable=self.auto_scroll_var
        )
        self.auto_scroll_check.pack(side="left", padx=5, pady=5)
        
        # 清除按鈕
        self.clear_button = ctk.CTkButton(
            self.toolbar,
            text="🗑 清除",
            width=80,
            command=self.clear
        )
        self.clear_button.pack(side="right", padx=(5, 10), pady=5)
        
        # 匯出按鈕
        self.export_button = ctk.CTkButton(
            self.toolbar,
            text="📤 匯出",
            width=80,
            command=self.export_logs
        )
        self.export_button.pack(side="right", padx=5, pady=5)
        
        # 日誌文本區域
        self.textbox = ctk.CTkTextbox(self)
        self.textbox.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.logs = []  # 儲存所有日誌
    
    def add_log(self, level: str, message: str, timestamp: str = None):
        """添加日誌"""
        if not timestamp:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        log_entry = {
            "timestamp": timestamp,
            "level": level,
            "message": message
        }
        self.logs.append(log_entry)
        
        # 如果符合過濾條件，顯示日誌
        if self._should_show_log(log_entry):
            self._display_log(log_entry)
        
        # 自動滾動到底部
        if self.auto_scroll_var.get():
            self.textbox.see("end")
    
    def _should_show_log(self, log_entry: dict) -> bool:
        """檢查是否應該顯示此日誌"""
        level_filter = self.level_var.get()
        return level_filter == "ALL" or log_entry["level"] == level_filter
    
    def _display_log(self, log_entry: dict):
        """顯示日誌條目"""
        level_colors = {
            "INFO": "green",
            "WARNING": "orange", 
            "ERROR": "red"
        }
        
        color = level_colors.get(log_entry["level"], "white")
        log_line = f"[{log_entry['timestamp']}] {log_entry['level']}: {log_entry['message']}\n"
        
        self.textbox.insert("end", log_line)
    
    def _filter_logs(self, selected_level: str):
        """過濾日誌顯示"""
        self.textbox.delete("1.0", "end")
        
        for log_entry in self.logs:
            if self._should_show_log(log_entry):
                self._display_log(log_entry)
    
    def clear(self):
        """清除所有日誌"""
        self.logs.clear()
        self.textbox.delete("1.0", "end")
    
    def export_logs(self):
        """匯出日誌"""
        # 這裡可以實現日誌匯出功能
        content = self.textbox.get("1.0", "end-1c")
        print("匯出日誌內容:")
        print(content)


class DataTable(ctk.CTkFrame):
    """數據表格組件"""
    
    def __init__(self, parent, headers: List[str], **kwargs):
        super().__init__(parent, **kwargs)
        
        self.headers = headers
        self.data = []
        
        # 創建表格框架
        self.table_frame = ctk.CTkScrollableFrame(self)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 創建表頭
        self._create_headers()
    
    def _create_headers(self):
        """創建表頭"""
        for i, header in enumerate(self.headers):
            header_label = ctk.CTkLabel(
                self.table_frame,
                text=header,
                font=ctk.CTkFont(size=12, weight="bold"),
                fg_color="gray20",
                corner_radius=5
            )
            header_label.grid(row=0, column=i, padx=2, pady=2, sticky="ew")
            self.table_frame.grid_columnconfigure(i, weight=1)
    
    def add_row(self, row_data: List[str]):
        """添加行數據"""
        row_index = len(self.data) + 1
        self.data.append(row_data)
        
        for i, cell_data in enumerate(row_data):
            cell_label = ctk.CTkLabel(
                self.table_frame,
                text=str(cell_data),
                font=ctk.CTkFont(size=11),
                fg_color="transparent"
            )
            cell_label.grid(row=row_index, column=i, padx=2, pady=1, sticky="ew")
    
    def clear_data(self):
        """清除所有數據"""
        self.data.clear()
        # 清除除表頭外的所有widgets
        for widget in self.table_frame.winfo_children():
            if widget.grid_info()['row'] > 0:
                widget.destroy()
    
    def update_data(self, new_data: List[List[str]]):
        """更新表格數據"""
        self.clear_data()
        for row in new_data:
            self.add_row(row)


class ProgressDialog(ctk.CTkToplevel):
    """進度對話框"""
    
    def __init__(self, parent, title: str = "處理中", message: str = "請稍候..."):
        super().__init__(parent)
        
        self.title(title)
        self.geometry("400x150")
        self.resizable(False, False)
        
        # 設置為模態對話框
        self.transient(parent)
        self.grab_set()
        
        # 居中顯示
        self.center_window()
        
        # 訊息標籤
        self.message_label = ctk.CTkLabel(
            self,
            text=message,
            font=ctk.CTkFont(size=14)
        )
        self.message_label.pack(pady=20)
        
        # 進度條
        self.progress = ctk.CTkProgressBar(self, width=300)
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        # 取消按鈕
        self.cancel_button = ctk.CTkButton(
            self,
            text="取消",
            width=100,
            command=self.cancel
        )
        self.cancel_button.pack(pady=10)
        
        self.cancelled = False
    
    def center_window(self):
        """將窗口居中"""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.winfo_screenheight() // 2) - (150 // 2)
        self.geometry(f"400x150+{x}+{y}")
    
    def update_progress(self, value: float, message: str = None):
        """更新進度"""
        self.progress.set(value)
        if message:
            self.message_label.configure(text=message)
        self.update()
    
    def cancel(self):
        """取消操作"""
        self.cancelled = True
        self.destroy()
    
    def is_cancelled(self) -> bool:
        """檢查是否被取消"""
        return self.cancelled