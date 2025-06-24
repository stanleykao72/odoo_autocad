# utility/popup_selector.py

import tkinter as tk
from tkinter import ttk, messagebox

class PopupSelector:
    def __init__(self, parent, fetch_data_func, target_entry, columns, title="選擇一個選項", additional_target_entries=None):
        """
        :param parent: 父小部件
        :param fetch_data_func: 獲取數據的函數，應返回一個字典列表
        :param target_entry: 用於填充選定數據的 Entry 小部件
        :param columns: 要顯示的數據鍵列表
        :param title: 彈出窗口的標題
        :param additional_target_entries: 字典，映射額外欄位到其 Entry 小部件
        """
        self.parent = parent
        self.fetch_data_func = fetch_data_func
        self.target_entry = target_entry
        self.columns = columns
        self.title = title
        self.additional_target_entries = additional_target_entries or {}
        self.popup_open = False  # 新增：追蹤彈出窗口是否已開啟
        self.popup = None        # 新增：存儲彈出窗口實例
        self.prevent_show = False  # 新增：防止循環觸發

    def show(self):
        if self.popup_open or self.prevent_show:
            self.prevent_show = False  # 重置標誌
            return
        self.popup_open = True
        data = self.fetch_data_func()
        if not data:
            messagebox.showerror("錯誤", "未能獲取數據。")
            self.popup_open = False
            return

        self.popup = tk.Toplevel(self.parent)  # 修改：將 popup 定義為實例屬性
        self.popup.title(self.title)
        self.popup.geometry("400x300")
        self.popup.resizable(False, False)  # 可選：禁止改變彈出窗口大小

        # 將彈出視窗置中
        self._center_popup()

        # 綁定關閉事件
        self.popup.protocol("WM_DELETE_WINDOW", self.on_close)

        # 搜尋輸入框
        self.search_input = tk.Entry(self.popup)
        self.search_input.pack(fill="x", padx=10, pady=5)
        self.search_input.bind("<KeyRelease>", self.update_filter)

        # 資料顯示區域 (Treeview)
        self.tree = ttk.Treeview(self.popup, columns=self.columns, show='headings')
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        for item in data:
            values = [item.get(col, "") for col in self.columns]
            self.tree.insert('', tk.END, values=values)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)

        # 綁定雙擊事件
        self.tree.bind("<Double-1>", self.select_item)

        # 添加滾動條
        scrollbar = ttk.Scrollbar(self.popup, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y', pady=5)

    def _center_popup(self):
        self.parent.update_idletasks()
        self.popup.update_idletasks()

        # 獲取父視窗的位置和尺寸
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()

        # 獲取彈出視窗的尺寸
        popup_width = self.popup.winfo_width()
        popup_height = self.popup.winfo_height()

        # 計算彈出視窗的位置，使其置於父視窗中央
        x = parent_x + (parent_width - popup_width) // 2
        y = parent_y + (parent_height - popup_height) // 2

        # 設定彈出視窗的位置
        self.popup.geometry(f"+{x}+{y}")

    def update_filter(self, event):
        search_text = event.widget.get().lower()
        self.filtered_items = [
            item for item in self.fetch_data_func()
            if any(str(value).lower().find(search_text) != -1
                   for value in item.values())
        ]
        self._update_tree_items()

    def _update_tree_items(self):
        # 清空現有項目
        for i in self.tree.get_children():
            self.tree.delete(i)
        # 插入過濾後的項目
        for item in self.filtered_items:
            values = [item.get(col, "") for col in self.columns]
            self.tree.insert("", "end", values=values)

    def select_item(self, event):
        selected = self.tree.focus()
        if selected:
            values = self.tree.item(selected, "values")
            if values:
                self.target_entry.delete(0, tk.END)
                self.target_entry.insert(0, values[0])  # 填充主要欄位 (例如 'name')

                # 填充額外的 Entry 欄位
                for i, col in enumerate(self.columns[1:], start=1):
                    entry = self.additional_target_entries.get(col)
                    if entry and i < len(values):
                        entry.configure(state='normal')
                        entry.delete(0, tk.END)
                        entry.insert(0, values[i])
                        entry.configure(state='readonly')

                self.prevent_show = True  # 新增：防止下一次 FocusIn 調用 show
                self.on_close()  # 關閉彈出窗口

    def on_close(self):
        if self.popup:
            self.popup.destroy()
            self.popup = None
        self.popup_open = False
