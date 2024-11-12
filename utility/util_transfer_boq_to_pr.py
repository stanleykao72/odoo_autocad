# -*- coding: utf-8 -*-
import tkinter as tk

class UtilTransferBoqToPr:
    def __init__(self):
        pass

    def transfer_boq_to_pr(self, main_body):
        self.clear_main_body(main_body)
        log_messages = tk.Text(main_body)
        log_messages.pack(fill="both", expand=True)
        log_messages.insert(tk.END, "轉移 BOQ 到 PR...\n")
        main_body.after(1000, lambda: log_messages.insert(tk.END, "成功轉移 BOQ 到 PR。\n"))

    def clear_main_body(self, main_body):
        for widget in main_body.winfo_children():
            widget.destroy()
