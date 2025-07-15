# -*- coding: utf-8 -*-
"""
MCP Request Handler for AutoCAD-Odoo Integration

統一的MCP工具註冊和請求處理器
整合AutoCAD讀圖功能和Odoo資料操作
遵循CLAUDE.md的TDD方法論和依賴注入模式
"""
import json
import traceback
from typing import Dict, Any, Callable, Optional, List


class MCPRequestHandler:
    """
    MCP請求處理器
    
    負責註冊MCP工具、處理JSON-RPC請求，並執行相應的AutoCAD或Odoo操作
    """
    
    def __init__(self, autocad_util, odoo_util, log_util):
        """
        初始化MCP請求處理器
        
        Args:
            autocad_util: AutoCAD工具實例
            odoo_util: Odoo工具實例  
            log_util: 日誌工具實例
        """
        # 依賴注入 (遵循CLAUDE.md模式)
        self.autocad_util = autocad_util
        self.odoo_util = odoo_util
        self.log_util = log_util
        
        # 工具註冊表
        self.tools: Dict[str, Dict[str, Any]] = {}
        
        # 自動註冊所有內建工具
        self._register_builtin_tools()
    
    def register_tool(self, name: str, function: Callable, description: str, 
                     parameters: Optional[Dict] = None) -> None:
        """
        註冊MCP工具
        
        Args:
            name: 工具名稱
            function: 工具執行函數
            description: 工具描述
            parameters: 工具參數定義 (JSON Schema格式)
        """
        self.tools[name] = {
            "function": function,
            "description": description,
            "parameters": parameters or {}
        }
        
        self.log_util.safe_log_insert(f"MCP工具已註冊: {name}\n")
    
    def execute_tool(self, tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """
        執行指定的MCP工具
        
        Args:
            tool_name: 工具名稱
            arguments: 工具參數
            
        Returns:
            Dict: 工具執行結果
        """
        if tool_name not in self.tools:
            raise ValueError(f"Unknown tool: {tool_name}")
        
        try:
            tool_function = self.tools[tool_name]["function"]
            result = tool_function(**arguments)
            
            self.log_util.safe_log_insert(f"MCP工具執行成功: {tool_name}\n")
            return result
            
        except Exception as e:
            error_msg = f"MCP工具執行失敗 {tool_name}: {str(e)}"
            self.log_util.safe_log_insert(f"{error_msg}\n")
            raise RuntimeError(error_msg)
    
    def process_request(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """
        處理JSON-RPC MCP請求
        
        Args:
            request: JSON-RPC請求字典
            
        Returns:
            Dict: JSON-RPC回應字典
        """
        try:
            # 驗證請求格式
            if not self._validate_request(request):
                return self._create_error_response(
                    request.get("id"), -32600, "Invalid Request"
                )
            
            method = request["method"]
            params = request.get("params", {})
            request_id = request["id"]
            
            # 處理不同類型的MCP請求
            if method == "tools/list":
                return self._handle_tools_list(request_id)
            elif method == "tools/call":
                return self._handle_tool_call(request_id, params)
            else:
                return self._create_error_response(
                    request_id, -32601, f"Method not found: {method}"
                )
                
        except Exception as e:
            self.log_util.safe_log_insert(f"MCP請求處理錯誤: {str(e)}\n")
            return self._create_error_response(
                request.get("id"), -32603, "Internal error"
            )
    
    def _register_builtin_tools(self) -> None:
        """註冊所有內建MCP工具"""
        # AutoCAD讀圖工具
        self._register_autocad_tools()
        
        # Odoo整合工具
        self._register_odoo_tools()
    
    def _register_autocad_tools(self) -> None:
        """註冊AutoCAD相關的MCP工具"""
        
        def scan_all_entities() -> Dict[str, Any]:
            """掃描所有AutoCAD實體"""
            try:
                entities = self.autocad_util.scan_entities()
                return {
                    "success": True,
                    "entities": entities,
                    "count": len(entities) if entities else 0
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        def get_table_data(layout: str = "Model") -> Dict[str, Any]:
            """讀取表格資料"""
            try:
                table_data = self.autocad_util.get_layouts_values()
                return {
                    "success": True,
                    "tables": table_data.get("tables", []),
                    "layout": layout
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        def extract_layout_info() -> Dict[str, Any]:
            """提取圖面佈局資訊"""
            try:
                # 這裡會整合現有的autocad_util功能
                layouts = getattr(self.autocad_util, 'get_layouts', lambda: [])()
                return {
                    "success": True,
                    "layouts": layouts
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        def query_entities_by_type(entity_type: str) -> Dict[str, Any]:
            """按類型查詢實體"""
            try:
                # 先掃描所有實體，然後篩選
                all_entities = self.autocad_util.scan_entities()
                filtered_entities = [
                    entity for entity in all_entities 
                    if entity.get('type', '').upper() == entity_type.upper()
                ]
                return {
                    "success": True,
                    "entities": filtered_entities,
                    "entity_type": entity_type,
                    "count": len(filtered_entities)
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        # 註冊AutoCAD工具
        self.register_tool(
            "scan_all_entities", 
            scan_all_entities,
            "掃描並分析AutoCAD圖面中的所有實體"
        )
        
        self.register_tool(
            "get_table_data",
            get_table_data, 
            "讀取AutoCAD圖面中的表格資料",
            {"layout": {"type": "string", "default": "Model"}}
        )
        
        self.register_tool(
            "extract_layout_info",
            extract_layout_info,
            "提取AutoCAD圖面的佈局資訊"
        )
        
        self.register_tool(
            "query_entities_by_type",
            query_entities_by_type,
            "按實體類型查詢AutoCAD圖面元素",
            {"entity_type": {"type": "string", "required": True}}
        )
    
    def _register_odoo_tools(self) -> None:
        """註冊Odoo相關的MCP工具"""
        
        def get_products_by_query(query: str, limit: int = 50) -> Dict[str, Any]:
            """按查詢條件搜尋Odoo產品"""
            try:
                products = self.odoo_util.search_products(query, limit)
                return {
                    "success": True,
                    "products": products,
                    "query": query,
                    "count": len(products) if products else 0
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        def push_boq_to_project(project_id: int, boq_data: List[Dict]) -> Dict[str, Any]:
            """推送BOQ資料到Odoo專案"""
            try:
                # 這裡會整合現有的push_to_boq功能
                result = self.odoo_util.push_boq_data(project_id, boq_data)
                return {
                    "success": True,
                    "project_id": project_id,
                    "items_pushed": len(boq_data),
                    "result": result
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        def validate_product_mapping(product_codes: List[str]) -> Dict[str, Any]:
            """驗證產品代碼對應"""
            try:
                validation_results = []
                for code in product_codes:
                    # 檢查產品是否存在於Odoo中
                    products = self.odoo_util.search_products(code, 1)
                    validation_results.append({
                        "code": code,
                        "exists": len(products) > 0,
                        "product": products[0] if products else None
                    })
                
                return {
                    "success": True,
                    "validations": validation_results,
                    "total_checked": len(product_codes)
                }
            except Exception as e:
                return {"success": False, "error": str(e)}
        
        # 註冊Odoo工具
        self.register_tool(
            "get_products_by_query",
            get_products_by_query,
            "按查詢條件搜尋Odoo產品資料庫",
            {
                "query": {"type": "string", "required": True},
                "limit": {"type": "integer", "default": 50}
            }
        )
        
        self.register_tool(
            "push_boq_to_project",
            push_boq_to_project,
            "推送BOQ資料到指定的Odoo專案",
            {
                "project_id": {"type": "integer", "required": True},
                "boq_data": {"type": "array", "required": True}
            }
        )
        
        self.register_tool(
            "validate_product_mapping",
            validate_product_mapping,
            "驗證產品代碼是否存在於Odoo系統中",
            {
                "product_codes": {"type": "array", "required": True}
            }
        )
    
    def _validate_request(self, request: Dict[str, Any]) -> bool:
        """驗證JSON-RPC請求格式"""
        required_fields = ["jsonrpc", "method", "id"]
        return all(field in request for field in required_fields)
    
    def _handle_tools_list(self, request_id: Any) -> Dict[str, Any]:
        """處理工具列表請求"""
        tools_list = []
        for name, info in self.tools.items():
            tools_list.append({
                "name": name,
                "description": info["description"],
                "inputSchema": {
                    "type": "object",
                    "properties": info.get("parameters", {}),
                }
            })
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "tools": tools_list
            }
        }
    
    def _handle_tool_call(self, request_id: Any, params: Dict[str, Any]) -> Dict[str, Any]:
        """處理工具調用請求"""
        try:
            tool_name = params["name"]
            arguments = params.get("arguments", {})
            
            result = self.execute_tool(tool_name, arguments)
            
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps(result, ensure_ascii=False, indent=2)
                        }
                    ]
                }
            }
            
        except Exception as e:
            return self._create_error_response(
                request_id, -32603, f"Tool execution failed: {str(e)}"
            )
    
    def _create_error_response(self, request_id: Any, code: int, message: str) -> Dict[str, Any]:
        """建立錯誤回應"""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "error": {
                "code": code,
                "message": message
            }
        }