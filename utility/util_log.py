# util_log.py

import os
import logging
from datetime import datetime

import tkinter as tk
from tkinter import scrolledtext

# File logger setup
_file_logger = logging.getLogger("odoo_autocad")
_file_logger.setLevel(logging.DEBUG)

# Log directory: logs/ under project root
_log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")
os.makedirs(_log_dir, exist_ok=True)

# Daily log file: logs/odoo_autocad_YYYY-MM-DD.log
_log_file = os.path.join(_log_dir, f"odoo_autocad_{datetime.now().strftime('%Y-%m-%d')}.log")
_file_handler = logging.FileHandler(_log_file, encoding="utf-8")
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
_file_logger.addHandler(_file_handler)


class UtilLog:
    def __init__(self, parent):
        self.parent = parent
        if parent is not None:
            # GUI模式：建立ScrolledText組件
            self.log_messages = scrolledtext.ScrolledText(parent, height=10, bg="black", fg="white", state='disabled')
            self.log_messages.pack(fill="both", expand=True)
            self.gui_mode = True
        else:
            # 無GUI模式：使用控制台輸出
            self.log_messages = None
            self.gui_mode = False

        # Log version at session start
        try:
            from version import APP_VERSION
        except ImportError:
            APP_VERSION = "unknown"
        self.safe_log_insert(f"Log Initialized. (v{APP_VERSION})\n")
        _file_logger.info(f"=== Session started v{APP_VERSION} (log file: {_log_file}) ===")

    def safe_log_insert(self, message):
        """安全地插入日誌訊息，同時寫入檔案和GUI。"""
        # Always write to file
        _file_logger.info(message.strip())

        if self.gui_mode and self.log_messages and self.log_messages.winfo_exists():
            # GUI模式：更新ScrolledText組件
            self.log_messages.configure(state='normal')  # 啟用編輯
            self.log_messages.insert(tk.END, message)
            self.log_messages.see(tk.END)  # 滾動到最後
            self.log_messages.configure(state='disabled')  # 禁用編輯
        else:
            # 無GUI模式或組件不存在：輸出到控制台
            print(f"[LOG] {message.strip()}")
