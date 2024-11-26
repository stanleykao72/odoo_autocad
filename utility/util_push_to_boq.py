# -*- coding: utf-8 -*-


class UtilPushToBoq:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def push_to_boq(self):

        layout_dic = self.autocad_util.get_layouts_values()
        self.log_util.safe_log_insert(f"layout_dic: {layout_dic}\n")

        boq_list = self.odoo_util.import2boq(layout_dic)
        self.log_util.safe_log_insert(f"boq_list: {boq_list}\n")

        self.autocad_util.set_layouts_tables_id(boq_list)

        # self.log_util.safe_log_insert(f"return_table_list: {return_table_list}\n")
        self.log_util.safe_log_insert("推送到 BOQ...\n")
