# util_log.py

import tkinter as tk
from tkinter import scrolledtext

class UtilLog:
    def __init__(self, parent):
        self.log_messages = scrolledtext.ScrolledText(parent, height=10, bg="black", fg="white", state='disabled')
        self.log_messages.pack(fill="both", expand=True)
        self.safe_log_insert("Log Initialized.\n")

    def safe_log_insert(self, message):
        """安全地插入日誌訊息，避免因小部件無效而引發錯誤。"""
        if self.log_messages and self.log_messages.winfo_exists():
            self.log_messages.configure(state='normal')  # 啟用編輯
            self.log_messages.insert(tk.END, message)
            self.log_messages.see(tk.END)  # 滾動到最後
            self.log_messages.configure(state='disabled')  # 禁用編輯
        else:
            print("Log Messages widget does not exist. Message:", message)

    # def insert_message(self, message):
    #     """
    #     安全地插入日誌訊息，確保線程安全。

    #     :param message: 要插入的訊息。
    #     """
    #     if self.log_messages and self.log_messages.winfo_exists():
    #         # 使用 after 方法在主線程中更新 UI
    #         self.log_messages.after(0, self._insert, message)
    #     else:
    #         print("Log Messages widget does not exist. Message:", message)

    # def _insert(self, message):
    #     """
    #     實際插入訊息的方法，應在主線程中執行。

    #     :param message: 要插入的訊息。
    #     """
    #     self.log_messages.configure(state='normal')  # 啟用編輯
    #     self.log_messages.insert(tk.END, message)
    #     self.log_messages.see(tk.END)  # 滾動到最後
    #     self.log_messages.configure(state='disabled')  # 禁用編輯