# -*- coding: utf-8 -*-
import tkinter as tk

class UtilTransferBoqToPr:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def transfer_boq_to_pr(self):
        self.log_util.safe_log_insert("轉移 BOQ 到 PR...\n")

