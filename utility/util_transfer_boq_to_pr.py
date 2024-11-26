# -*- coding: utf-8 -*-
import tkinter as tk

class UtilTransferBoqToPr:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def transfer_boq_to_pr(self):
        header_id_dict = self.autocad_util.get_layouts_header_id_to_pr()
        self.log_util.safe_log_insert(f"header_id_dict: {header_id_dict}\n")

        pr_list = self.odoo_util.boq2pr(header_id_dict)

        self.log_util.safe_log_insert(f"pr_list: {pr_list}\n")

        self.log_util.safe_log_insert("轉移 BOQ 到 PR...\n")

