# -*- coding: utf-8 -*-
"""
AutoCAD Dispatcher — COM / IPC 雙模式統一介面

根據使用者選擇的模式，將呼叫路由到 UtilAutoCAD (COM) 或 UtilAutoCADIPC (File IPC)。
設計參考 C# 端的 DrawingDataServiceDispatcher。
"""

import logging

_logger = logging.getLogger(__name__)


class UtilAutoCADDispatcher:
    """COM/IPC 模式切換器，提供與 UtilAutoCAD 相同的介面"""

    MODE_COM = "com"
    MODE_IPC = "ipc"

    def __init__(self, odoo_util, log_util, mode="com"):
        self.odoo_util = odoo_util
        self.log_util = log_util
        self._mode = mode
        self._com_backend = None
        self._ipc_backend = None
        self._init_backend(mode)

    def _init_backend(self, mode):
        """Initialize the backend for the given mode"""
        if mode == self.MODE_COM:
            if self._com_backend is None:
                from utility.util_autocad import UtilAutoCAD
                self._com_backend = UtilAutoCAD(self.odoo_util, self.log_util)
            self._active = self._com_backend
        elif mode == self.MODE_IPC:
            if self._ipc_backend is None:
                from utility.util_autocad_ipc import UtilAutoCADIPC
                self._ipc_backend = UtilAutoCADIPC(log_util=self.log_util)
            self._active = self._ipc_backend
        else:
            raise ValueError(f"Unknown AutoCAD mode: {mode}")

    @property
    def mode(self):
        return self._mode

    def switch_mode(self, new_mode):
        """Switch between COM and IPC modes"""
        if new_mode == self._mode:
            return
        old_mode = self._mode
        self._mode = new_mode
        self._init_backend(new_mode)
        self._log(f"[Dispatcher] Mode switched: {old_mode} → {new_mode}\n")

    @property
    def active_backend(self):
        return self._active

    def _log(self, msg):
        if self.log_util:
            self.log_util.safe_log_insert(msg)
        else:
            _logger.info(msg.rstrip('\n'))

    # === Proxy properties (forward to active backend) ===

    @property
    def acad(self):
        return getattr(self._active, 'acad', None)

    @property
    def project_id(self):
        return getattr(self._active, 'project_id', None)

    @project_id.setter
    def project_id(self, value):
        self._active.project_id = value

    @property
    def project_name(self):
        return getattr(self._active, 'project_name', None)

    @project_name.setter
    def project_name(self, value):
        self._active.project_name = value

    @property
    def layout_name(self):
        return getattr(self._active, 'layout_name', None)

    @property
    def pr_no(self):
        return getattr(self._active, 'pr_no', None)

    @pr_no.setter
    def pr_no(self, value):
        self._active.pr_no = value

    @property
    def job_working_plan_id(self):
        return getattr(self._active, 'job_working_plan_id', None)

    @job_working_plan_id.setter
    def job_working_plan_id(self, value):
        self._active.job_working_plan_id = value

    @property
    def job_working_plan_name(self):
        return getattr(self._active, 'job_working_plan_name', None)

    @job_working_plan_name.setter
    def job_working_plan_name(self, value):
        self._active.job_working_plan_name = value

    # === Connection ===

    def connected_autocad(self):
        return self._active.connected_autocad()

    def connect_autocad(self, main_body=None):
        self._active.connect_autocad(main_body)
        # IPC mode: read block attributes and look up project from Odoo
        # (COM mode does this internally in UtilAutoCAD.connect_autocad)
        if self._mode == self.MODE_IPC and self._active.connected_autocad():
            self._process_pr_no_ipc()

    # === Layout Management ===

    def get_active_layout(self):
        return self._active.get_active_layout()

    def get_doc_layouts(self):
        return self._active.get_doc_layouts()

    # === Odoo Operations ===

    def get_layouts_values(self):
        return self._active.get_layouts_values()

    def get_single_layout_values(self, layout_name):
        if hasattr(self._active, 'get_single_layout_values'):
            return self._active.get_single_layout_values(layout_name)
        # COM fallback: not supported, caller should use get_layouts_values
        return {}

    def set_layouts_tables_id(self, boq_list):
        return self._active.set_layouts_tables_id(boq_list)

    def get_layouts_header_id_to_pr(self):
        return self._active.get_layouts_header_id_to_pr()

    def get_block_attributes(self):
        return self._active.get_block_attributes()

    def set_block_attributes(self, attrs, layout_name=None):
        return self._active.set_block_attributes(attrs, layout_name)

    def clear_table_id(self, layout=None):
        return self._active.clear_table_id(layout)

    def clear_all_tables_id(self):
        return self._active.clear_all_tables_id()

    # === Drawing Operations ===

    def draw_line(self, start_point, end_point, layer="0"):
        return self._active.draw_line(start_point, end_point, layer)

    def draw_circle(self, center_point, radius, layer="0"):
        return self._active.draw_circle(center_point, radius, layer)

    def set_layer(self, layer_name, color=7, create_if_not_exist=True):
        return self._active.set_layer(layer_name, color, create_if_not_exist)

    def list_layers(self, filter_type="all", sort_by="name", include_details=True):
        return self._active.list_layers(filter_type, sort_by, include_details)

    def scan_elements(self, element_type="all", **kwargs):
        return self._active.scan_elements(element_type, **kwargs)

    # === IPC project initialization ===

    def _process_pr_no_ipc(self):
        """Read pr_no + layout_name from block attributes via IPC, look up project from Odoo"""
        try:
            attrs = self._active.get_block_attributes()
            self._log(f"[IPC] Block attributes: {attrs}\n")
            if not attrs:
                self._log("[IPC] No attribute block found — ensure 050_block_util.lsp is reloaded in AutoCAD\n")
                return

            # Extract layout_name (added by ob:action-get-block-attrs)
            layout_name = attrs.get('layout_name', '')
            if layout_name:
                self._active.layout_name = layout_name
                self._log(f"[IPC] Layout: {layout_name}\n")

            pr_no = attrs.get('pr_no', '')
            if not pr_no:
                self._log(f"[IPC] No 'pr_no' tag in block (found tags: {list(attrs.keys())})\n")
                return

            self._active.pr_no = pr_no
            self._log(f"[IPC] PR No: {pr_no}\n")

            # Look up project from Odoo
            if self.odoo_util:
                project = self.odoo_util.get_project(pr_no)
                if project:
                    self._active.project_id = project.get('id')
                    self._active.project_name = project.get('name')
                    self._active.job_working_plan_id = project.get('job_working_plan_id')
                    self._active.job_working_plan_name = project.get('job_working_plan_name')
                    self._log(f"[IPC] Project: {self._active.project_name} (ID: {self._active.project_id})\n")
                else:
                    self._log(f"[IPC] No project found for PR No: {pr_no}\n")
            else:
                self._log("[IPC] Odoo not connected, cannot look up project\n")
        except Exception as e:
            self._log(f"[IPC] process_pr_no failed: {e}\n")

    # === COM-only methods (graceful fallback for IPC) ===

    def clear_main_body(self, main_body):
        if hasattr(self._active, 'clear_main_body'):
            return self._active.clear_main_body(main_body)

    def get_layout_attribute_blocks_value(self, layout, tag_list):
        if hasattr(self._active, 'get_layout_attribute_blocks_value'):
            return self._active.get_layout_attribute_blocks_value(layout, tag_list)
        return {}

    def get_layout_table_block(self, blocks):
        if hasattr(self._active, 'get_layout_table_block'):
            return self._active.get_layout_table_block(blocks)
        return []

    def get_table_data(self, block_list):
        if hasattr(self._active, 'get_table_data'):
            return self._active.get_table_data(block_list)
        return None, []

    def get_block_text(self, block):
        if hasattr(self._active, 'get_block_text'):
            return self._active.get_block_text(block)
        return {}
