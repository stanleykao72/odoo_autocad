# -*- coding: utf-8 -*-


class UtilPushToBoq:
    def __init__(self, odoo_util, autocad_util, log_util):
        self.odoo_util = odoo_util
        self.autocad_util = autocad_util
        self.log_util = log_util

    def push_to_boq(self):
        table_list = self.autocad_util.get_tables_from_layouts()
        # self.log_util.safe_log_insert(f"table_list: {table_list}\n")
        return_table_list = []
        for table in table_list:
            self.log_util.safe_log_insert(f"table: {table}\n")
            table_dict = self.autocad_util.get_table_data(table['layout'], table['block'])
            return_table_list.append(table_dict)
        self.log_util.safe_log_insert(f"return_table_list: {return_table_list}\n")
        self.log_util.safe_log_insert("推送到 BOQ...\n")
