# -*- coding: utf-8 -*-
"""
AutoCAD Dispatcher — COM / IPC 雙模式統一介面

根據使用者選擇的模式，將呼叫路由到 UtilAutoCAD (COM) 或 UtilAutoCADIPC (File IPC)。
設計參考 C# 端的 DrawingDataServiceDispatcher。
"""

import logging
import threading

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
        # Version selection: COM progid or IPC target HWND
        self._com_progid = "AutoCAD.Application"
        self._ipc_target_hwnd = None
        # COM 跨線程支援（見 _invoke）
        self._gui_proxy = None
        self._gui_thread_id = None
        self._proxy_timeout = 120.0
        self._init_backend(mode)

    # === COM 跨線程路由 ===

    def set_gui_proxy(self, proxy, gui_thread_id, timeout=None):
        """註冊 GUI 代理，讓 COM 模式下的跨線程呼叫能繞回 GUI 主線程。

        COM 物件是 STA 封送的：在 GUI 主線程建立、從背景線程（例如 MCP
        server 的 uvicorn thread）呼叫，會得到 RPC_E_WRONG_THREAD
        (0x8001010E)。而 UtilAutoCAD 內部多處 try/except 會把它吞掉並回傳
        空結果 —— 表現成「圖面沒有資料」，極難診斷。

        註冊後，COM 模式下非主線程的呼叫會自動改走代理；主線程與 IPC 模式
        則維持直接呼叫（IPC 走檔案與視窗訊息，沒有線程限制）。
        """
        self._gui_proxy = proxy
        self._gui_thread_id = gui_thread_id
        if timeout is not None:
            self._proxy_timeout = timeout
        self._log(f"[Dispatcher] GUI 代理已註冊（主線程 id={gui_thread_id}）\n")

    def _needs_proxy(self):
        """是否需要把呼叫繞回 GUI 主線程"""
        return (self._mode == self.MODE_COM
                and self._gui_proxy is not None
                and self._gui_thread_id is not None
                and threading.get_ident() != self._gui_thread_id)

    def _invoke(self, method, *args, **kwargs):
        """呼叫後端方法；必要時透過 GUI 代理在主線程執行"""
        if not self._needs_proxy():
            return getattr(self._active, method)(*args, **kwargs)

        response = self._gui_proxy.execute_in_gui(
            "call_backend", timeout=self._proxy_timeout,
            method=method, args=args, kwargs=kwargs)

        if response.get("success"):
            return response.get("result")

        error = response.get("error", "GUI 代理執行失敗")
        self._log(f"[Dispatcher] {method} 經 GUI 代理失敗: {error}\n")
        raise RuntimeError(f"{method}: {error}")

    def set_autocad_version(self, progid=None, target_hwnd=None):
        """Set target AutoCAD version for next connection.
        Args:
            progid: COM ProgID (e.g. "AutoCAD.Application.19" for 2014)
            target_hwnd: IPC window HWND for specific AutoCAD instance
        """
        if progid:
            self._com_progid = progid
            if self._com_backend:
                self._com_backend.autocad_progid = progid
        if target_hwnd:
            self._ipc_target_hwnd = target_hwnd
            if self._ipc_backend:
                self._ipc_backend._target_hwnd = target_hwnd

    def _init_backend(self, mode):
        """Initialize the backend for the given mode"""
        if mode == self.MODE_COM:
            if self._com_backend is None:
                from utility.util_autocad import UtilAutoCAD
                self._com_backend = UtilAutoCAD(self.odoo_util, self.log_util)
                self._com_backend.autocad_progid = self._com_progid
            self._active = self._com_backend
        elif mode == self.MODE_IPC:
            if self._ipc_backend is None:
                from utility.util_autocad_ipc import UtilAutoCADIPC
                self._ipc_backend = UtilAutoCADIPC(
                    log_util=self.log_util, target_hwnd=self._ipc_target_hwnd)
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
        # COM 端只檢查 self.acad / self.doc 是否為 None，不實際呼叫 COM，
        # 從任何線程判斷都安全；不繞代理可避免 MCP 狀態查詢卡在代理往返上。
        return self._active.connected_autocad()

    def connect_autocad(self, main_body=None):
        self._active.connect_autocad(main_body)
        # IPC mode: read block attributes and look up project from Odoo
        # (COM mode does this internally in UtilAutoCAD.connect_autocad)
        if self._mode == self.MODE_IPC and self._active.connected_autocad():
            self._process_pr_no_ipc()

    # === Layout Management ===
    #
    # 以下方法都會碰到 AutoCAD 的資料，一律經 _invoke —— COM 模式下由背景
    # 線程呼叫時會自動繞回 GUI 主線程（見 set_gui_proxy）。

    def get_active_layout(self):
        return self._invoke('get_active_layout')

    def get_doc_layouts(self):
        return self._invoke('get_doc_layouts')

    # === Odoo Operations ===

    def get_layouts_values(self):
        return self._invoke('get_layouts_values')

    def get_single_layout_values(self, layout_name):
        return self._invoke('get_single_layout_values', layout_name)

    def set_layouts_tables_id(self, boq_list):
        return self._invoke('set_layouts_tables_id', boq_list)

    def get_layouts_header_id_to_pr(self):
        return self._invoke('get_layouts_header_id_to_pr')

    def get_block_attributes(self):
        return self._invoke('get_block_attributes')

    def set_block_attributes(self, attrs, layout_name=None):
        return self._invoke('set_block_attributes', attrs, layout_name)

    def clear_table_id(self, layout=None):
        return self._invoke('clear_table_id', layout)

    def clear_all_tables_id(self):
        return self._invoke('clear_all_tables_id')

    # === Drawing Operations ===

    def draw_line(self, start_point, end_point, layer="0", layout_name=None):
        return self._invoke('draw_line', start_point, end_point, layer, layout_name)

    def draw_circle(self, center_point, radius, layer="0", layout_name=None):
        return self._invoke('draw_circle', center_point, radius, layer, layout_name)

    def set_layer(self, layer_name, color=7, create_if_not_exist=True):
        return self._invoke('set_layer', layer_name, color, create_if_not_exist)

    def list_layers(self, filter_type="all", sort_by="name", include_details=True):
        return self._invoke('list_layers', filter_type, sort_by, include_details)

    def scan_elements(self, element_type="all", **kwargs):
        return self._invoke('scan_elements', element_type, **kwargs)

    # === Layout change detection & refresh ===

    def refresh_active_layout(self):
        """Re-read current active layout and project info.

        Returns the new layout name if it changed, else None.
        Used by the GUI polling loop to detect layout switches.

        COM mode: get_active_layout() is cheap (in-process COM).
        IPC mode: uses single get_block_attributes() call to minimise
                  IPC round-trips (returns layout_name + pr_no in one shot).
        """
        try:
            old_layout = getattr(self._active, 'layout_name', None)

            if self._mode == self.MODE_IPC:
                return self._refresh_active_layout_ipc(old_layout)

            # --- COM mode ---
            new_layout = self.get_active_layout()
            if new_layout and not isinstance(new_layout, str):
                new_layout = getattr(new_layout, 'Name', str(new_layout))

            if not new_layout or new_layout == old_layout:
                return None

            self._active.layout_name = new_layout
            self._log(f"[Dispatcher] Layout changed: {old_layout} → {new_layout}\n")
            self._active.process_pr_no(new_layout)
            return new_layout
        except Exception as e:
            self._log(f"[Dispatcher] refresh_active_layout failed: {e}\n")
            return None

    def _refresh_active_layout_ipc(self, old_layout):
        """IPC-optimised refresh: single get_block_attributes() call
        returns layout_name + pr_no, avoiding multiple IPC round-trips."""
        attrs = self._active.get_block_attributes()
        if not attrs:
            return None

        new_layout = attrs.get('layout_name', '')
        if not new_layout or new_layout == old_layout:
            return None

        self._active.layout_name = new_layout
        self._log(f"[Dispatcher] Layout changed: {old_layout} → {new_layout}\n")

        # pr_no + project lookup (reuse attrs already fetched)
        pr_no = attrs.get('pr_no', '')
        if pr_no:
            self._active.pr_no = pr_no
            self._log(f"[IPC] PR No: {pr_no}\n")
            if self.odoo_util:
                project = self.odoo_util.get_project(pr_no)
                if project:
                    odoo_project_name = project.get('name', '')
                    odoo_jwp_name = project.get('job_working_plan_name', '')

                    self._active.project_id = project.get('id')
                    self._active.project_name = odoo_project_name
                    self._active.job_working_plan_id = project.get('job_working_plan_id')
                    self._active.job_working_plan_name = odoo_jwp_name
                    self._log(f"[IPC] Project: {self._active.project_name}\n")

                    # Compare with drawing values — update ALL layouts if different
                    self._sync_project_attrs_if_differ(attrs, odoo_project_name, odoo_jwp_name)

        return new_layout

    # === IPC project initialization ===

    def _process_pr_no_ipc(self):
        """Read pr_no + layout_name from block attributes via IPC, look up project from Odoo.

        When Odoo returns project_name or job_working_plan_name that differ from
        the values already stored in the drawing, update ALL layouts at once
        (set_block_attributes without layout_name → AutoLISP writes all layouts).
        """
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
                    odoo_project_name = project.get('name', '')
                    odoo_jwp_name = project.get('job_working_plan_name', '')

                    self._active.project_id = project.get('id')
                    self._active.project_name = odoo_project_name
                    self._active.job_working_plan_id = project.get('job_working_plan_id')
                    self._active.job_working_plan_name = odoo_jwp_name
                    self._log(f"[IPC] Project: {self._active.project_name} (ID: {self._active.project_id})\n")

                    # Compare with drawing values — update ALL layouts if different
                    self._sync_project_attrs_if_differ(attrs, odoo_project_name, odoo_jwp_name)
                else:
                    self._log(f"[IPC] No project found for PR No: {pr_no}\n")
            else:
                self._log("[IPC] Odoo not connected, cannot look up project\n")
        except Exception as e:
            self._log(f"[IPC] process_pr_no failed: {e}\n")

    def _sync_project_attrs_if_differ(self, drawing_attrs, odoo_project_name, odoo_jwp_name,
                                       layout_name=None):
        """Compare Odoo project_name / job_working_plan_name with drawing values.

        Args:
            drawing_attrs: block attribute dict from current layout
            odoo_project_name: project name from Odoo
            odoo_jwp_name: job_working_plan_name from Odoo
            layout_name: if provided, update only this layout;
                         if None, update ALL layouts at once.
        """
        dwg_project_name = drawing_attrs.get('project_name', '')
        dwg_jwp_name = drawing_attrs.get('job_working_plan_name', '')

        need_update = (
            (odoo_project_name and odoo_project_name != dwg_project_name)
            or (odoo_jwp_name and odoo_jwp_name != dwg_jwp_name)
        )
        if not need_update:
            return

        diff_parts = []
        if odoo_project_name != dwg_project_name:
            diff_parts.append(f"project_name: '{dwg_project_name}' → '{odoo_project_name}'")
        if odoo_jwp_name != dwg_jwp_name:
            diff_parts.append(f"job_working_plan_name: '{dwg_jwp_name}' → '{odoo_jwp_name}'")

        target = layout_name if layout_name else "所有 Layout"
        self._log(f"[IPC] Odoo 與圖面不同，更新 {target}: {'; '.join(diff_parts)}\n")

        update_attrs = {
            'project_name': odoo_project_name,
            'job_working_plan_name': odoo_jwp_name,
        }
        self._active.set_block_attributes(update_attrs, layout_name=layout_name)
        self._log(f"[IPC] 已同步 project_name / job_working_plan_name 至 {target}\n")

    # === COM-only methods (forwarded; only available when COM backend is active) ===

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
