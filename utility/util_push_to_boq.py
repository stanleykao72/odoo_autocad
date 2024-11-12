# -*- coding: utf-8 -*-
import tkinter as tk

class UtilPushToBoq:
    def __init__(self):
        pass

    def push_to_boq(self, main_body):
        self.clear_main_body(main_body)
        log_messages = tk.Text(main_body)
        log_messages.pack(fill="both", expand=True)
        log_messages.insert(tk.END, "推送到 BOQ...\n")
        main_body.after(1000, lambda: log_messages.insert(tk.END, "成功推送到 BOQ。\n"))

    def clear_main_body(self, main_body):
        for widget in main_body.winfo_children():
            widget.destroy()
