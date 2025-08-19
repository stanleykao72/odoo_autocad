# -*- coding: utf-8 -*-
"""
MCP SSE Server Manager for GUI Integration
Directly integrates MCPSSEServer class instead of subprocess management
"""

import threading
import asyncio
import time
import logging
import requests
import uvicorn
from typing import Optional, Dict, Any
import sys
import os

# Add parent directory to path to import modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# Import the standard MCP SSE server
from utility.util_mcp_sse_server import StandardMCPSSEServer

class MCPSSEManager:
    """Manager for MCP SSE Server in GUI using Standard MCP Protocol"""
    
    def __init__(self, port: int = 8084, autocad_util=None, odoo_util=None):
        self.port = port
        self.server: Optional[StandardMCPSSEServer] = None
        self.is_running = False
        self.status_callback = None
        
        # Store shared utility instances
        self.autocad_util = autocad_util
        self.odoo_util = odoo_util
        
        # 快取AutoCAD狀態資訊（由主執行緒更新）
        self.autocad_status_cache = {
            "document_name": None,
            "version": None,
            "layout_info": None,
            "last_update": None
        }
        
        # Configure logging
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        self.logger.info(f"[SSE Manager] MCPSSEManager 初始化，端口: {port} (標準 MCP 協定)")
        
        # Initialize the standard MCP server
        self._initialize_standard_server()
        
        # 設置共享的 AutoCAD 實例給 MCP Server (AC6 要求)
        if self.autocad_util:
            try:
                # Import the MCP server to set shared instance
                import mcp_server_fastmcp
                mcp_server_fastmcp.set_shared_autocad_util(self.autocad_util)
                self.logger.info("[SSE Manager] AutoCAD 共享實例已設置給 MCP Server")
            except Exception as e:
                self.logger.warning(f"[SSE Manager] 設置 AutoCAD 共享實例失敗: {e}")
        
        # 設置共享的 Odoo 實例給 MCP Server
        if self.odoo_util:
            try:
                import mcp_server_fastmcp
                mcp_server_fastmcp.set_shared_odoo_util(self.odoo_util)
                self.logger.info("[SSE Manager] Odoo 共享實例已設置給 MCP Server")
            except Exception as e:
                self.logger.warning(f"[SSE Manager] 設置 Odoo 共享實例失敗: {e}")
        
        # 初始化時嘗試更新快取
        self.update_autocad_status_cache()
    
    def update_autocad_status_cache(self):
        """更新AutoCAD狀態快取（應在主執行緒中調用）"""
        # 確保 MCP Server 使用最新的共享 AutoCAD 實例 (AC6 要求)
        if self.autocad_util:
            try:
                import mcp_server_fastmcp
                mcp_server_fastmcp.set_shared_autocad_util(self.autocad_util)
                self.logger.debug("[SSE Manager] AutoCAD 共享實例已更新")
            except Exception as e:
                self.logger.warning(f"[SSE Manager] 更新 AutoCAD 共享實例失敗: {e}")
        
        try:
            import time
            if self.autocad_util and self.autocad_util.connected_autocad():
                # 嘗試獲取文檔名稱
                doc_name = None
                try:
                    if hasattr(self.autocad_util, 'doc') and self.autocad_util.doc:
                        if hasattr(self.autocad_util.doc, 'Name'):
                            doc_name = str(self.autocad_util.doc.Name)
                except Exception:
                    pass
                
                # 嘗試獲取版本資訊
                version = None
                try:
                    if hasattr(self.autocad_util, 'acad') and self.autocad_util.acad:
                        if hasattr(self.autocad_util.acad, 'Version'):
                            version = str(self.autocad_util.acad.Version)
                except Exception:
                    pass
                
                # 嘗試獲取layout資訊
                layout_info = None
                try:
                    if hasattr(self.autocad_util, 'acad') and self.autocad_util.acad:
                        # 使用更簡單的get_doc_layouts方法獲取真實layouts
                        try:
                            layouts_list = self.autocad_util.get_doc_layouts()
                            if layouts_list:
                                all_layouts = []
                                current_layout_name = None
                                
                                # 嘗試獲取當前活動layout
                                current_layout_name = None
                                try:
                                    # 方法1: 嘗試使用get_active_layout()方法
                                    active_layout = self.autocad_util.get_active_layout()
                                    if active_layout and hasattr(active_layout, 'Name'):
                                        current_layout_name = str(active_layout.Name)
                                        self.logger.debug(f"透過get_active_layout()檢測到當前layout: {current_layout_name}")
                                except Exception as e1:
                                    self.logger.debug(f"get_active_layout()方法失敗: {e1}")
                                    try:
                                        # 方法2: 嘗試直接存取ActiveDocument.ActiveLayout
                                        if hasattr(self.autocad_util.acad, 'ActiveDocument') and self.autocad_util.acad.ActiveDocument:
                                            if hasattr(self.autocad_util.acad.ActiveDocument, 'ActiveLayout'):
                                                current_layout_name = str(self.autocad_util.acad.ActiveDocument.ActiveLayout.Name)
                                                self.logger.debug(f"透過ActiveDocument.ActiveLayout檢測到當前layout: {current_layout_name}")
                                    except Exception as e2:
                                        self.logger.debug(f"ActiveDocument.ActiveLayout方法失敗: {e2}")
                                        # 最終回退：使用第一個非Model的layout作為當前layout
                                        pass
                                
                                # 遍歷所有layouts
                                for i, layout in enumerate(layouts_list):
                                    try:
                                        layout_name = str(layout.Name) if hasattr(layout, 'Name') else f"Layout{i}"
                                        tab_order = getattr(layout, 'TabOrder', i) if hasattr(layout, 'TabOrder') else i
                                        
                                        all_layouts.append({
                                            "name": layout_name,
                                            "tab_order": tab_order,
                                            "is_current": layout_name == current_layout_name
                                        })
                                    except Exception:
                                        # 如果無法獲取layout屬性，跳過
                                        continue
                                
                                if all_layouts:
                                    current_layout = None
                                    
                                    # 尋找已標記為current的layout
                                    for layout in all_layouts:
                                        if layout["is_current"]:
                                            current_layout = layout
                                            break
                                    
                                    # 如果沒有找到current layout，使用智慧回退策略
                                    if not current_layout and all_layouts:
                                        # 優先選擇非"Model"的第一個layout作為當前layout
                                        for layout in all_layouts:
                                            if layout["name"] != "Model":
                                                layout["is_current"] = True
                                                current_layout = layout
                                                self.logger.debug(f"回退策略：設定 '{layout['name']}' 為當前layout")
                                                break
                                        
                                        # 如果所有layout都是"Model"或沒有找到合適的，使用第一個
                                        if not current_layout:
                                            all_layouts[0]["is_current"] = True
                                            current_layout = all_layouts[0]
                                            self.logger.debug(f"最終回退：設定 '{current_layout['name']}' 為當前layout")
                                    
                                    layout_info = {
                                        "current_layout": current_layout,
                                        "all_layouts": all_layouts,
                                        "total_layouts": len(all_layouts)
                                    }
                        except Exception as layout_ex:
                            self.logger.debug(f"get_doc_layouts failed: {layout_ex}")
                except Exception as e:
                    self.logger.debug(f"獲取layout資訊時發生錯誤: {e}")
                
                # 更新快取
                self.autocad_status_cache.update({
                    "document_name": doc_name,
                    "version": version,
                    "layout_info": layout_info,
                    "last_update": time.time()
                })
                
        except Exception as e:
            self.logger.debug(f"[SSE Manager] 更新AutoCAD狀態快取失敗: {e}")
    
    def _initialize_standard_server(self):
        """Initialize the standard MCP SSE server"""
        try:
            self.server = StandardMCPSSEServer(port=self.port, host="localhost")
            
            # Register additional tools with utility instances
            if self.autocad_util:
                self._register_autocad_tools()
                self.logger.info("[SSE Manager] AutoCAD 工具已註冊到 MCP 伺服器")
            
            if self.odoo_util:
                self._register_odoo_tools()
                self.logger.info("[SSE Manager] Odoo 工具已註冊到 MCP 伺服器")
                
        except Exception as e:
            self.logger.error(f"[SSE Manager] 初始化標準 MCP 伺服器時發生錯誤: {e}")
    
    def _register_autocad_tools(self):
        """Register AutoCAD-specific tools"""
        if not self.server or not self.autocad_util:
            return
            
        def extract_autocad_parameters(drawing_path: str = "", use_current_drawing: bool = True):
            """Extract parameters from AutoCAD drawing (通過GUI代理執行)"""
            try:
                if not self.autocad_util.connected_autocad():
                    return {"success": False, "error": "AutoCAD 未連接"}
                
                # 通過GUI代理執行參數提取
                from utility.util_gui_proxy import get_gui_proxy
                gui_proxy = get_gui_proxy()
                
                result = gui_proxy.execute_in_gui(
                    "extract_parameters", 
                    drawing_path=drawing_path, 
                    use_current_drawing=use_current_drawing
                )
                
                if result.get("success"):
                    # 重新格式化為原有格式
                    parameters = result.get("parameters", [])
                    return {
                        "success": True,
                        "parameters": parameters,
                        "drawing_info": {
                            "document": "Current Drawing",
                            "total_layouts": len(parameters)
                        },
                        "extraction_time": "Real-time",
                        "method": "GUI代理執行"
                    }
                else:
                    return {
                        "success": False,
                        "error": result.get("error", "參數提取失敗"),
                        "method": "GUI代理執行"
                    }
                    
            except Exception as e:
                return {"success": False, "error": str(e), "method": "GUI代理執行"}
        
        self.server.register_tool(
            name="extract_autocad_parameters",
            description="從 AutoCAD 圖檔提取參數資訊",
            schema={
                "type": "object",
                "properties": {
                    "drawing_path": {"type": "string", "description": "圖檔路徑 (可選)"},
                    "use_current_drawing": {"type": "boolean", "description": "使用當前開啟的圖檔", "default": True}
                },
                "required": []
            },
            function=extract_autocad_parameters
        )
        
        # 添加 AutoCAD 狀態檢查工具，使用共享的 AutoCAD 實例
        def check_autocad_status():
            """Check AutoCAD connection status (通過GUI代理獲取詳細狀態)"""
            try:
                # 先檢查基本連接
                is_connected = self.autocad_util.connected_autocad()
                
                if not is_connected:
                    return {
                        "connected": False,
                        "status": "AutoCAD 未連接",
                        "suggestion": "請啟動 AutoCAD 並嘗試重新連接"
                    }
                
                # 通過GUI代理獲取詳細狀態
                from utility.util_gui_proxy import get_gui_proxy
                gui_proxy = get_gui_proxy()
                
                result = gui_proxy.execute_in_gui("get_autocad_status")
                
                if result.get("success"):
                    # 添加方法資訊
                    result["method"] = "GUI代理執行"
                    return result
                else:
                    # GUI代理失敗，使用快取資訊
                    return {
                        "connected": True,
                        "status": "連接但無法獲取詳細資訊",
                        "error": result.get("error", "GUI代理失敗"),
                        "method": "GUI代理執行(失敗)"
                    }
                    
            except Exception as e:
                return {
                    "connected": False,
                    "status": "檢查失敗",
                    "error": str(e),
                    "method": "GUI代理執行"
                }
        
        self.server.register_tool(
            name="check_autocad_status",
            description="檢查 AutoCAD 連接狀態 (使用共享實例)",
            schema={
                "type": "object",
                "properties": {},
                "required": []
            },
            function=check_autocad_status
        )
    
    def _register_odoo_tools(self):
        """Register Odoo-specific tools"""
        if not self.server or not self.odoo_util:
            return
            
        def check_odoo_status():
            """Check Odoo connection status"""
            try:
                # Use the connected_odoo method that exists in UtilOdoo
                is_connected = self.odoo_util.connected_odoo()
                
                if is_connected and self.odoo_util.odoo:
                    # Try to get additional info from the odoo client
                    try:
                        server_info = {}
                        if hasattr(self.odoo_util.odoo, 'swagger_spec') and self.odoo_util.odoo.swagger_spec:
                            server_info["server_url"] = self.odoo_util.odoo.swagger_spec.origin_url
                        
                        return {
                            "connected": True,
                            "server_url": server_info.get("server_url", "Connected"),
                            "database": "Connected Database",
                            "user_info": {"status": "authenticated"},
                            "status": "已連接"
                        }
                    except Exception:
                        # Fallback to basic connected status
                        return {
                            "connected": True,
                            "server_url": "Connected",
                            "database": "Connected Database", 
                            "user_info": {"status": "authenticated"},
                            "status": "已連接"
                        }
                else:
                    return {
                        "connected": False,
                        "server_url": "N/A",
                        "database": "N/A",
                        "user_info": {},
                        "status": "連接失敗"
                    }
            except Exception as e:
                return {"connected": False, "error": str(e), "status": "連接錯誤"}
        
        self.server.register_tool(
            name="check_odoo_status",
            description="檢查 Odoo 伺服器連接狀態",
            schema={
                "type": "object",
                "properties": {},
                "required": []
            },
            function=check_odoo_status
        )
        
        # Layout 相關工具
        def get_current_layout():
            """獲取當前活動的 layout (通過GUI代理獲取實時資料)"""
            try:
                # 先檢查AutoCAD連接
                if not self.autocad_util.connected_autocad():
                    return {"success": False, "error": "AutoCAD 未連接"}
                
                # 通過GUI代理獲取實時current layout
                from utility.util_gui_proxy import get_gui_proxy
                gui_proxy = get_gui_proxy()
                
                current_layout_result = gui_proxy.execute_in_gui("get_current_layout")
                
                if not current_layout_result.get("success"):
                    return current_layout_result
                
                # 只返回當前layout資訊，不需要列出所有layouts
                result = {
                    "success": True,
                    "current_layout": current_layout_result["current_layout"],  # 實時資料
                    "data_source": "realtime_gui_proxy",
                    "method": "GUI代理執行"
                }
                
                return result
                    
            except Exception as e:
                return {"success": False, "error": str(e), "method": "GUI代理執行"}
        
        def switch_to_layout(layout_name: str):
            """切換到指定的 layout"""
            try:
                if not self.autocad_util.connected_autocad():
                    return {"success": False, "error": "AutoCAD 未連接"}
                
                # 記錄當前狀態用於診斷
                original_layout = None
                try:
                    original_layout = self.autocad_util.get_active_layout()
                    self.logger.info(f"當前 layout: {original_layout.Name if original_layout else 'None'}")
                except Exception as e:
                    self.logger.warning(f"無法獲取當前 layout: {e}")
                
                # 使用快取的 layout 資訊避免 COM 線程問題 (AC6 修正)
                available_layouts = []
                target_layout_name = None
                
                # 從 layout 快取中獲取資訊，避免直接 COM 調用
                try:
                    # 先嘗試從快取獲取
                    layout_info = self.autocad_status_cache.get("layout_info")
                    if layout_info and "all_layouts" in layout_info:
                        for layout_data in layout_info["all_layouts"]:
                            layout_name_clean = layout_data["name"].strip()
                            available_layouts.append(layout_name_clean)
                            
                            # 嘗試多種匹配方式
                            if (layout_name_clean == layout_name or 
                                layout_name_clean.lower() == layout_name.lower()):
                                target_layout_name = layout_data["name"]
                                self.logger.info(f"從快取中找到匹配的 layout: '{target_layout_name}' (搜尋: '{layout_name}')")
                                break
                    
                    # 如果快取中沒有找到，嘗試直接檢查（在主線程中執行）
                    if not target_layout_name and self.autocad_util:
                        self.logger.info("快取中未找到 layout，嘗試通過共享實例查找")
                        # 這裡不直接訪問 COM 對象，而是返回錯誤讓用戶知道需要更新快取
                        if not available_layouts:
                            return {
                                "success": False, 
                                "error": f"找不到名為 '{layout_name}' 的 layout，請先使用 get_current_layout 更新 layout 快取",
                                "suggestion": "請先調用 get_current_layout 工具來更新 layout 資訊快取"
                            }
                        
                except Exception as e:
                    self.logger.error(f"獲取 layout 資訊時發生錯誤: {e}")
                    return {"success": False, "error": f"獲取 layout 資訊失敗: {str(e)}"}
                
                if not target_layout_name:
                    return {
                        "success": False, 
                        "error": f"找不到名為 '{layout_name}' 的 layout",
                        "available_layouts": available_layouts,
                        "search_attempted": layout_name,
                        "debug_info": f"搜尋了 {len(available_layouts)} 個 layouts",
                        "suggestion": "請檢查 layout 名稱是否正確，或使用 get_current_layout 更新快取"
                    }
                
                # 使用GUI代理執行layout切換 (解決COM線程問題)
                try:
                    self.logger.info(f"通過GUI代理切換到 layout: {target_layout_name}")
                    
                    # 導入GUI代理
                    from utility.util_gui_proxy import get_gui_proxy
                    gui_proxy = get_gui_proxy()
                    
                    # 通過GUI線程執行切換
                    result = gui_proxy.execute_in_gui("switch_layout", layout_name=target_layout_name)
                    
                    if result.get("success"):
                        self.logger.info(f"GUI代理切換成功: {target_layout_name}")
                        return {
                            "success": True,
                            "message": result.get("message", f"成功切換到 layout: {target_layout_name}"),
                            "current_layout": result.get("current_layout"),
                            "method": "GUI代理執行"
                        }
                    else:
                        error_msg = result.get("error", "GUI代理執行失敗")
                        self.logger.error(f"GUI代理切換失敗: {error_msg}")
                        return {
                            "success": False, 
                            "error": f"無法切換到 layout '{layout_name}'",
                            "error_details": [f"GUI代理執行失敗: {error_msg}"],
                            "attempted_methods": ["GUI代理執行"],
                            "suggestion": "請確認 AutoCAD 連接正常且 layout 名稱正確"
                        }
                    
                    # 等待一小段時間讓 AutoCAD 完成切換
                    import time
                    time.sleep(0.1)
                    
                    # 確認切換成功
                    current_layout = self.autocad_util.get_active_layout()
                    if current_layout and current_layout.Name == layout_name:
                        self.logger.info(f"切換成功驗證通過: {current_layout.Name}")
                        return {
                            "success": True,
                            "message": f"成功切換到 layout: {layout_name}",
                            "current_layout": {
                                "name": current_layout.Name,
                                "tab_order": current_layout.TabOrder if hasattr(current_layout, 'TabOrder') else 0,
                                "is_model_space": current_layout.Name == "Model"
                            }
                        }
                    else:
                        current_name = current_layout.Name if current_layout else "None"
                        self.logger.warning(f"切換後驗證失敗: 期望={layout_name}, 實際={current_name}")
                        return {
                            "success": False, 
                            "error": f"切換 layout 後驗證失敗，期望: '{layout_name}', 實際: '{current_name}'",
                            "debug_info": {
                                "target_layout_found": target_layout.Name,
                                "current_after_switch": current_name,
                                "original_layout": original_layout.Name if original_layout else "None"
                            }
                        }
                        
                except Exception as switch_error:
                    self.logger.error(f"切換過程中發生異常: {switch_error}")
                    return {"success": False, "error": f"切換 layout 時發生錯誤: {str(switch_error)}"}
                    
            except Exception as e:
                self.logger.error(f"switch_to_layout 總體異常: {e}")
                return {"success": False, "error": f"函數執行失敗: {str(e)}"}
        
        def draw_in_layout(layout_name: str, shape_type: str, parameters: dict):
            """在指定的 layout 中繪圖 (通過GUI代理執行)"""
            try:
                if not self.autocad_util.connected_autocad():
                    return {"success": False, "error": "AutoCAD 未連接"}
                
                # 通過GUI代理獲取當前layout
                from utility.util_gui_proxy import get_gui_proxy
                gui_proxy = get_gui_proxy()
                
                current_result = gui_proxy.execute_in_gui("get_current_layout")
                if not current_result.get("success"):
                    return current_result
                
                original_name = current_result["current_layout"]["name"]
                
                # 切換到目標 layout (如果需要)
                if layout_name != original_name:
                    switch_result = switch_to_layout(layout_name)
                    if not switch_result.get("success"):
                        return switch_result
                
                try:
                    # 通過GUI代理根據形狀類型繪圖
                    if shape_type == "line":
                        start_point = parameters.get("start_point", [0, 0, 0])
                        end_point = parameters.get("end_point", [100, 100, 0])
                        layer = parameters.get("layer", "0")
                        
                        result = gui_proxy.execute_in_gui(
                            "draw_line", 
                            start_point=start_point, 
                            end_point=end_point, 
                            layer=layer
                        )
                        draw_result = {
                            "shape_type": "line",
                            "parameters": {
                                "start_point": start_point,
                                "end_point": end_point,
                                "layer": layer
                            },
                            "result": result,
                            "method": "GUI代理執行"
                        }
                        
                    elif shape_type == "circle":
                        center_point = parameters.get("center_point", [0, 0, 0])
                        radius = parameters.get("radius", 10)
                        layer = parameters.get("layer", "0")
                        
                        result = gui_proxy.execute_in_gui(
                            "draw_circle",
                            center_point=center_point,
                            radius=radius,
                            layer=layer
                        )
                        draw_result = {
                            "shape_type": "circle",
                            "parameters": {
                                "center_point": center_point,
                                "radius": radius,
                                "layer": layer
                            },
                            "autocad_result": result
                        }
                        
                    else:
                        return {
                            "success": False,
                            "error": f"不支援的圖形類型: {shape_type}",
                            "supported_types": ["line", "circle"]
                        }
                    
                    return {
                        "success": True,
                        "message": f"成功在 layout '{layout_name}' 中繪製 {shape_type}",
                        "layout_name": layout_name,
                        "draw_result": draw_result
                    }
                    
                finally:
                    # 恢復原來的 layout (如果有切換的話)
                    if layout_name != original_name and original_name:
                        try:
                            switch_to_layout(original_name)
                        except Exception:
                            pass  # 忽略恢復錯誤
                        
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        def extract_layout_parameters(layout_name: str):
            """提取指定layout的AutoCAD參數 (通過GUI代理執行)"""
            try:
                if not self.autocad_util.connected_autocad():
                    return {"success": False, "error": "AutoCAD 未連接"}
                
                # 通過GUI代理執行layout特定參數提取
                from utility.util_gui_proxy import get_gui_proxy
                gui_proxy = get_gui_proxy()
                
                result = gui_proxy.execute_in_gui("extract_layout_parameters", layout_name=layout_name)
                
                if result.get("success"):
                    return {
                        "success": True,
                        "layout_name": result.get("layout_name"),
                        "parameters": result.get("parameters", {}),        # 標題區塊資料
                        "table_data": result.get("table_data", {}),        # 表格資料
                        "extraction_method": result.get("extraction_method", "single_layout"),
                        "method": "GUI代理執行"
                    }
                else:
                    return {
                        "success": False,
                        "error": result.get("error", "layout參數提取失敗"),
                        "method": "GUI代理執行"
                    }
                    
            except Exception as e:
                return {"success": False, "error": str(e), "method": "GUI代理執行"}
        
        # 註冊 layout 相關工具
        self.server.register_tool(
            name="get_current_layout",
            description="獲取當前活動的 AutoCAD layout 及所有可用的 layouts",
            schema={
                "type": "object",
                "properties": {},
                "required": []
            },
            function=get_current_layout
        )
        
        self.server.register_tool(
            name="extract_layout_parameters",
            description="提取指定 layout 的 AutoCAD 參數資訊",
            schema={
                "type": "object",
                "properties": {
                    "layout_name": {
                        "type": "string",
                        "description": "要提取參數的 layout 名稱，例如 'S405-203'"
                    }
                },
                "required": ["layout_name"]
            },
            function=extract_layout_parameters
        )
        
        self.server.register_tool(
            name="switch_to_layout",
            description="切換到指定名稱的 AutoCAD layout",
            schema={
                "type": "object",
                "properties": {
                    "layout_name": {
                        "type": "string",
                        "description": "要切換到的 layout 名稱，例如 'Layout1', 'Sheet1', 'Model' 等"
                    }
                },
                "required": ["layout_name"]
            },
            function=switch_to_layout
        )
        
        self.server.register_tool(
            name="draw_in_layout",
            description="在指定的 layout 中繪製圖形 (支援線條和圓形)",
            schema={
                "type": "object",
                "properties": {
                    "layout_name": {
                        "type": "string",
                        "description": "要繪圖的 layout 名稱"
                    },
                    "shape_type": {
                        "type": "string", 
                        "description": "圖形類型",
                        "enum": ["line", "circle"]
                    },
                    "parameters": {
                        "type": "object",
                        "description": "圖形參數",
                        "properties": {
                            "start_point": {
                                "type": "array",
                                "description": "線條起點座標 [x, y, z]",
                                "items": {"type": "number"}
                            },
                            "end_point": {
                                "type": "array", 
                                "description": "線條終點座標 [x, y, z]",
                                "items": {"type": "number"}
                            },
                            "center_point": {
                                "type": "array",
                                "description": "圓心座標 [x, y, z]", 
                                "items": {"type": "number"}
                            },
                            "radius": {
                                "type": "number",
                                "description": "圓形半徑"
                            },
                            "layer": {
                                "type": "string",
                                "description": "圖層名稱，預設為 '0'",
                                "default": "0"
                            }
                        }
                    }
                },
                "required": ["layout_name", "shape_type", "parameters"]
            },
            function=draw_in_layout
        )
        
        def export_layout_image(layout_name: str, export_path: str = None, image_format: str = "wmf"):
            """匯出指定layout為圖像檔案"""
            try:
                if not self.autocad_util.connected_autocad():
                    return {"success": False, "error": "AutoCAD 未連接"}
                
                # 通過GUI代理執行圖像匯出
                from utility.util_gui_proxy import get_gui_proxy
                gui_proxy = get_gui_proxy()
                
                result = gui_proxy.execute_in_gui(
                    "export_layout_image", 
                    layout_name=layout_name,
                    export_path=export_path,
                    image_format=image_format
                )
                
                if result.get("success"):
                    return {
                        "success": True,
                        "message": result.get("message"),
                        "export_path": result.get("export_path"),
                        "file_size": result.get("file_size"),
                        "image_format": result.get("image_format"),
                        "layout_name": result.get("layout_name"),
                        "method": "GUI代理執行"
                    }
                else:
                    return {
                        "success": False,
                        "error": result.get("error", "圖像匯出失敗"),
                        "method": "GUI代理執行"
                    }
                    
            except Exception as e:
                return {"success": False, "error": str(e), "method": "GUI代理執行"}
        
        self.server.register_tool(
            name="export_layout_image",
            description="匯出指定 layout 為圖像檔案 (支援 WMF, BMP 等格式)",
            schema={
                "type": "object",
                "properties": {
                    "layout_name": {
                        "type": "string",
                        "description": "要匯出的 layout 名稱，例如 'S405-203'"
                    },
                    "export_path": {
                        "type": "string",
                        "description": "匯出檔案路徑 (可選，預設會在當前目錄生成)"
                    },
                    "image_format": {
                        "type": "string",
                        "description": "圖像格式",
                        "enum": ["wmf", "bmp", "eps"],
                        "default": "wmf"
                    }
                },
                "required": ["layout_name"]
            },
            function=export_layout_image
        )
        
    def set_status_callback(self, callback):
        """Set callback function for status updates"""
        self.status_callback = callback
        
    def _update_status(self, message: str, is_running: bool = None):
        """Update status and call callback if set"""
        old_status = self.is_running
        if is_running is not None:
            self.is_running = is_running
            
        self.logger.info(f"[SSE Manager] 狀態更新: {message} (運行狀態: {old_status} -> {self.is_running})")
        
        if self.status_callback:
            self.logger.debug(f"[SSE Manager] 呼叫狀態回調函數")
            try:
                self.status_callback(message, self.is_running)
                self.logger.debug(f"[SSE Manager] 狀態回調函數執行成功")
            except Exception as e:
                self.logger.error(f"[SSE Manager] 狀態回調函數執行失敗: {e}")
        else:
            self.logger.debug(f"[SSE Manager] 無狀態回調函數")
    
    def start_server(self) -> bool:
        """Start the standard MCP SSE server"""
        if self.is_running:
            self.logger.warning("[SSE Manager] 伺服器已在運行中")
            return True
            
        if not self.server:
            self.logger.error("[SSE Manager] 標準 MCP 伺服器未初始化")
            return False
            
        try:
            self.logger.info(f"[SSE Manager] 啟動標準 MCP SSE 伺服器 (端口: {self.port})")
            
            success = self.server.start_server()
            if success:
                self.is_running = True
                self._update_status(f"標準 MCP SSE 伺服器已啟動 (端口: {self.port})", True)
                self.logger.info("[SSE Manager] ✅ 伺服器啟動成功")
                return True
            else:
                self._update_status("啟動 MCP SSE 伺服器失敗", False)
                self.logger.error("[SSE Manager] ❌ 伺服器啟動失敗")
                return False
                
        except Exception as e:
            self.logger.error(f"[SSE Manager] 啟動伺服器時發生異常: {e}")
            self._update_status(f"啟動伺服器時發生錯誤: {str(e)}", False)
            return False
    
    def stop_server(self) -> bool:
        """Stop the standard MCP SSE server"""
        if not self.is_running:
            return True
            
        if not self.server:
            return True
            
        try:
            self.logger.info("[SSE Manager] 停止標準 MCP SSE 伺服器")
            
            success = self.server.stop_server()
            if success:
                self.is_running = False
                self._update_status("標準 MCP SSE 伺服器已停止", False)
                self.logger.info("[SSE Manager] ✅ 伺服器停止成功")
                return True
            else:
                self.logger.error("[SSE Manager] ❌ 伺服器停止失敗")
                return False
                
        except Exception as e:
            self.logger.error(f"[SSE Manager] 停止伺服器時發生異常: {e}")
            return False
    
    def restart_server(self) -> bool:
        """Restart the SSE server"""
        self._update_status("正在重新啟動 SSE 伺服器...")
        self.stop_server()
        time.sleep(1)
        return self.start_server()
    
    
    def get_server_status(self) -> Dict[str, Any]:
        """Get current server status"""
        if self.server:
            return self.server.get_server_status()
        else:
            return {
                "is_running": self.is_running,
                "port": self.port,
                "error": "Server not initialized"
            }
    
    def test_mcp_connection(self) -> Dict[str, Any]:
        """Test MCP connection to the server"""
        if self.server:
            return self.server.test_mcp_connection() 
        else:
            return {"success": False, "error": "標準 MCP 伺服器未初始化"}
    
    def cleanup(self):
        """Clean up resources"""
        if self.is_running:
            self.stop_server()