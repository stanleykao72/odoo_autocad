# -*- coding: utf-8 -*-
import logging
import base64
import requests
import time
import tkinter as tk
from tkinter import ttk, messagebox

from utility.util_odoo import UtilOdoo
from utility.util_autocad import UtilAutoCAD
from utility.util_push_to_boq import UtilPushToBoq
from utility.util_transfer_boq_to_pr import UtilTransferBoqToPr
from utility.util_log import UtilLog
from utility.util_mcp_sse_manager import MCPSSEManager
from forms.form_autocad_param import FormAutoCADParam
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

# _logger = logging.getLogger(__name__)
# _logger.setLevel(logging.INFO)

# # log to console
# c_handler = logging.StreamHandler()

# console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
# c_handler.setFormatter(console_format)
# c_handler.setLevel = logging.DEBUG

# _logger.addHandler(c_handler)

class FormMain(tk.Tk):
    def __init__(self, odoo_connection):
        super().__init__()
        self.odoo_connection = odoo_connection
        self.title("AutoCAD Odoo Integration")

        # 獲取屏幕寬度
        screen_width = self.winfo_screenwidth()
        window_width = 800
        window_height = 600

        # 計算窗口應該放置的位置
        x_position = screen_width - window_width
        y_position = 0

        # 設置窗口大小並將其放置在屏幕的最右邊
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Define main layout frames
        self.top_bar = tk.Frame(self, bg="orange", height=50)
        self.top_bar.pack(side="top", fill="x")

        self.side_menu = tk.Frame(self, width=200, bg="lightgray")
        self.side_menu.pack(side="left", fill="y")

        self.main_content = tk.Frame(self, bg="white")
        self.main_content.pack(side="top", fill="both", expand=True)

        self.log_frame = tk.Frame(self, bg="black")
        self.log_frame.pack(side="bottom", fill="x")

        # Initialize log utility
        self.log_util = UtilLog(self.log_frame)

        # Initialize Odoo and AutoCAD utilities
        self.odoo_util = UtilOdoo(self.odoo_connection, self.log_util)
        self.autocad_util = UtilAutoCAD(self.odoo_util, self.log_util)
        self.push_to_boq_util = UtilPushToBoq(self.odoo_util, self.autocad_util, self.log_util)
        self.transfer_boq_to_pr_util = UtilTransferBoqToPr(self.odoo_util, self.autocad_util, self.log_util)
        
        # Initialize MCP SSE Manager
        self.mcp_sse_manager = MCPSSEManager(port=8081)
        self.mcp_sse_manager.set_status_callback(self.on_mcp_status_update)

        # Add buttons to the side menu
        self.add_side_menu_buttons()

    def add_side_menu_buttons(self):
        buttons = [
            ("連接到 Odoo", self.connect_odoo, 'button_connect_odoo'),
            ("連接到 AutoCAD", self.connect_autocad, 'button_connect_autocad'),
            ("🤖 AI助手 (SSE)", self.toggle_mcp_sse, 'button_mcp_sse'),
            ("從 Odoo 獲取參數", self.get_parameters_from_odoo, 'button_get_parameters'),
            ("推送到 BOQ", self.push_to_boq, 'button_push_to_boq'),
            ("轉移 BOQ 到 PR", self.transfer_boq_to_pr, 'button_transfer_boq_to_pr'),
            ("清除此配置表格ID", self.clear_table_id, 'button_clear_table_id'),
            ("清除所有配置表格ID", self.clear_all_tables_id, 'button_clear_all_tables_id'),
        ]

        for (text, command, button_name) in buttons:
            button = tk.Button(self.side_menu, text=text, command=command, name=button_name)
            button.pack(fill="x", pady=5)
            setattr(self, button_name, button)  # 將按鈕賦值為類的屬性


        if self.odoo_util.connected_odoo():
            self.button_connect_odoo.config(bg="green")
        else:
            self.button_connect_odoo.config(bg="red")

        if self.autocad_util.connected_autocad():
            self.button_connect_autocad.config(bg="green")
        else:
            self.button_connect_autocad.config(bg="red")

    def connect_odoo(self):
        self.clear_main_content()
        self.log_util.safe_log_insert("連接到 Odoo...\n")
        try:
            odoo, requestOptions, token = self.odoo_util.connect_odoo(self.odoo_connection)
            # self.after(1000, lambda: self.log_util.safe_log_insert("與 Odoo 連線成功\n"))
            if self.odoo_util.connected_odoo():
                self.button_connect_odoo.config(bg="green")
        except requests.exceptions.ConnectionError:
            self.after(1000, lambda: self.log_util.safe_log_insert("無法與 Odoo 連線，通常多試幾次會成功\n"))
        except Exception as e:
            self.after(1000, lambda: self.log_util.safe_log_insert(f"連接 Odoo 時發生錯誤: {str(e)}\n"))

    def connect_autocad(self):
        self.autocad_util.connect_autocad(self.main_content)
        if self.autocad_util.connected_autocad():
            self.button_connect_autocad.config(bg="green")

    def get_parameters_from_odoo(self):
        # Ensure AutoCAD is connected and attributes are available
        if not self.autocad_util.project_id:
            self.show_error_message("錯誤", "請先連接 AutoCAD 並確保已獲取專案資料。")
            return

        # Initialize FormAutoCADParam with the UtilAutoCAD instance
        form_autocad_param = FormAutoCADParam(
            main_content=self.main_content,
            odoo_util=self.odoo_util,
            autocad_util=self.autocad_util,
            log_util=self.log_util,
            root=self
        )
        form_autocad_param.get_parameters_from_odoo()

    def push_to_boq(self):
        # Implement push to BOQ logic
        self.push_to_boq_util.push_to_boq()

    def transfer_boq_to_pr(self):
        # Implement transfer BOQ to PR logic
        self.transfer_boq_to_pr_util.transfer_boq_to_pr()
    
    def clear_table_id(self):
        self.autocad_util.clear_table_id()

    def clear_all_tables_id(self):
        self.autocad_util.clear_all_tables_id()

    def clear_main_content(self):
        for widget in self.main_content.winfo_children():
            widget.destroy()

    def show_error_message(self, title, message):
        self.update_idletasks()  # 確保窗口已更新
        messagebox.showerror(title, message, parent=self)
    
    def toggle_mcp_sse(self):
        """Toggle MCP SSE server on/off"""
        if self.mcp_sse_manager.is_running:
            self.stop_mcp_sse()
        else:
            self.start_mcp_sse()
    
    def start_mcp_sse(self):
        """Start MCP SSE server"""
        self.clear_main_content()
        
        # Create status display
        status_frame = tk.Frame(self.main_content)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(status_frame, text="🤖 AI助手 SSE 伺服器", font=("Arial", 14, "bold")).pack(anchor="w")
        
        self.mcp_status_label = tk.Label(status_frame, text="正在啟動...", fg="orange")
        self.mcp_status_label.pack(anchor="w")
        
        # Create control buttons
        control_frame = tk.Frame(self.main_content)
        control_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Button(control_frame, text="重新啟動", command=self.restart_mcp_sse).pack(side="left", padx=5)
        tk.Button(control_frame, text="測試連接", command=self.test_mcp_connection).pack(side="left", padx=5)
        tk.Button(control_frame, text="伺服器狀態", command=self.show_mcp_status).pack(side="left", padx=5)
        
        # Create log display
        log_frame = tk.Frame(self.main_content)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        tk.Label(log_frame, text="伺服器日誌:", font=("Arial", 10, "bold")).pack(anchor="w")
        
        self.mcp_log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        scrollbar = tk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.mcp_log_text.yview)
        self.mcp_log_text.configure(yscrollcommand=scrollbar.set)
        
        self.mcp_log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Start the server
        import threading
        threading.Thread(target=self._start_mcp_sse_async, daemon=True).start()
    
    def _start_mcp_sse_async(self):
        """Start MCP SSE server in background thread"""
        success = self.mcp_sse_manager.start_server()
        if success:
            self.after(0, self._update_mcp_button_success)
        else:
            self.after(0, self._update_mcp_button_error)
    
    def stop_mcp_sse(self):
        """Stop MCP SSE server"""
        self.mcp_sse_manager.stop_server()
        self._update_mcp_button_stopped()
    
    def restart_mcp_sse(self):
        """Restart MCP SSE server"""
        self.mcp_status_label.config(text="正在重新啟動...", fg="orange")
        import threading
        threading.Thread(target=self._restart_mcp_sse_async, daemon=True).start()
    
    def _restart_mcp_sse_async(self):
        """Restart MCP SSE server in background thread"""
        success = self.mcp_sse_manager.restart_server()
        if success:
            self.after(0, self._update_mcp_button_success)
        else:
            self.after(0, self._update_mcp_button_error)
    
    def test_mcp_connection(self):
        """Test MCP connection"""
        result = self.mcp_sse_manager.test_mcp_connection()
        
        if result["success"]:
            message = f"✅ 連接成功\\n工具數量: {result['tools_count']}\\n可用工具: {', '.join(result['tools'])}"
            messagebox.showinfo("MCP 連接測試", message)
        else:
            messagebox.showerror("MCP 連接測試", f"❌ 連接失敗\\n錯誤: {result['error']}")
    
    def show_mcp_status(self):
        """Show detailed MCP server status"""
        status = self.mcp_sse_manager.get_server_status()
        
        status_text = f"""
伺服器狀態: {'🟢 運行中' if status['is_running'] else '🔴 已停止'}
端口: {status['port']}
腳本: {status['server_script']}
健康檢查: {'✅ 正常' if status['health_check'] else '❌ 異常'}
        """
        
        messagebox.showinfo("MCP 伺服器狀態", status_text)
    
    def on_mcp_status_update(self, message: str, is_running: bool):
        """Callback for MCP status updates"""
        def update_ui():
            if hasattr(self, 'mcp_status_label'):
                color = "green" if is_running else "red"
                self.mcp_status_label.config(text=message, fg=color)
            
            if hasattr(self, 'mcp_log_text'):
                self.mcp_log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\\n")
                self.mcp_log_text.see(tk.END)
        
        self.after(0, update_ui)
    
    def _update_mcp_button_success(self):
        """Update MCP button to success state"""
        self.button_mcp_sse.config(text="🤖 AI助手 (運行中)", bg="green")
    
    def _update_mcp_button_error(self):
        """Update MCP button to error state"""
        self.button_mcp_sse.config(text="🤖 AI助手 (錯誤)", bg="red")
    
    def _update_mcp_button_stopped(self):
        """Update MCP button to stopped state"""
        self.button_mcp_sse.config(text="🤖 AI助手 (SSE)", bg="SystemButtonFace")
    
    def cleanup(self):
        """Clean up resources when closing"""
        if hasattr(self, 'mcp_sse_manager'):
            self.mcp_sse_manager.cleanup()

    def show_info_message(self, title, message):
        self.update_idletasks()  # 確保窗口已更新
        messagebox.showinfo(title, message, parent=self)

# class FormMain(tk.Tk):
#     def __init__(self, odoo=None, requestOptions=None, token=None, odoo_connection=None):
#         super().__init__()
#         self.odoo = odoo
#         self.requestOptions = requestOptions
#         self.token = token
#         self.odoo_connection = odoo_connection

#         self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
#         self.autocad_util = UtilAutoCAD(self.odoo_util)
#         self.push_to_boq_util = UtilPushToBoq()
#         self.transfer_boq_to_pr_util = UtilTransferBoqToPr()
#         self.filtered_items = []  # 初始化 filtered_items
        
#         # 設置視窗標題和大小
#         self.title("材料及加工方式")
#         window_width = 700
#         window_height = 450
        
#         # 獲取螢幕尺寸
#         screen_width = self.winfo_screenwidth()
#         screen_height = self.winfo_screenheight()
        
#         # 計算視窗位置 (靠右)
#         x = screen_width - window_width
#         y = (screen_height - window_height) // 2  # 垂直置中
        
#         # 設置視窗位置和大小
#         self.geometry(f"{window_width}x{window_height}+{x}+{y}")

#         # 初始化日誌訊息
#         self.autocad_util.initialize_log(self.log_frame)

#         self.create_widgets()

#     def create_widgets(self):
#         self.top_bar = tk.Frame(self, bg="orange", height=50)
#         self.top_bar.pack(side="top", fill="x")

#         self.side_menu = tk.Frame(self, bg="lightgrey", width=200)
#         self.side_menu.pack(side="left", fill="y")

#         self.main_body = tk.Frame(self, bg="white")
#         self.main_body.pack(side="right", fill="both", expand=True)

#         self.create_side_menu()

#     def create_side_menu(self):
#         self.menu_button_odoo = tk.Button(self.side_menu, text="連接到 Odoo", command=self.connect_odoo)
#         self.menu_button_odoo.pack(fill="x", pady=5)

#         self.menu_button_autocad = tk.Button(self.side_menu, text="連接到 AutoCAD", command=self.connect_autocad)
#         self.menu_button_autocad.pack(fill="x", pady=5)

#         self.menu_button_get_params = tk.Button(self.side_menu, text="從 Odoo 獲取參數", command=self.get_parameters_from_odoo)
#         self.menu_button_get_params.pack(fill="x", pady=5)

#         self.menu_button_push_boq = tk.Button(self.side_menu, text="推送到 BOQ", command=self.push_to_boq)
#         self.menu_button_push_boq.pack(fill="x", pady=5)

#         self.menu_button_transfer_boq = tk.Button(self.side_menu, text="轉移 BOQ 到 PR", command=self.transfer_boq_to_pr)
#         self.menu_button_transfer_boq.pack(fill="x", pady=5)

#     def connect_odoo(self):
#         self.clear_main_body()
#         self.log_messages = tk.Text(self.main_body)
#         self.log_messages.pack(fill="both", expand=True)
#         self.log_messages.insert(tk.END, "連接到 Odoo...\n")
#         try:
#             self.log_messages.insert(tk.END, f"self.odoo_connection: {self.odoo_connection}\n")
#             odoo, requestOptions, token = self.odoo_util.connect_odoo(self.odoo_connection)

#             log = self.log_messages
#             self.after(1000, lambda: log.insert(tk.END, "與 Odoo 連線成功\n"))
#             self.odoo = odoo
#             self.requestOptions = requestOptions
#             self.token = token
#             self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
#         except requests.exceptions.ConnectionError:
#             log = self.log_messages
#             self.after(1000, lambda: log.insert(tk.END, "無法與 Odoo 連線，通常多試幾次會成功\n"))
#         except Exception as e:
#             log = self.log_messages
#             self.after(1000, lambda: log.insert(tk.END, f"連接 Odoo 時發生錯誤: {str(e)}\n"))

#     def connect_autocad(self):
#         self.autocad_util.connect_autocad(self.main_body)

#     def get_parameters_from_odoo(self):
#         # Ensure AutoCAD is connected and attributes are available
#         if not self.autocad_util.project_id:
#             messagebox.showerror("錯誤", "請先連接 AutoCAD 並確保已獲取專案資料。")
#             return

#         # Extract required attributes
#         project_id = self.autocad_util.project_id

#         # Initialize FormAutoCADParam with extracted attributes
#         form_autocad_param = FormAutoCADParam(
#             main_body=self.main_body,
#             odoo_util=self.odoo_util,
#             util_autocad=self.autocad_util  # Pass the UtilAutoCAD instance
#         )
#         form_autocad_param.get_parameters_from_odoo()

#     def push_to_boq(self):
#         self.push_to_boq_util.push_to_boq(self.main_body)

#     def transfer_boq_to_pr(self):
#         self.transfer_boq_to_pr_util.transfer_boq_to_pr(self.main_body)

#     def clear_main_body(self):
#         for widget in self.main_body.winfo_children():
#             widget.destroy()

if __name__ == '__main__':
    odoo_connection = {
        'host': 'your_host',
        'db_name': 'your_db_name',
        'url': 'your_url',
        'token': 'your_token'
    }
    app = FormMain(odoo_connection=odoo_connection)
    app.mainloop()
