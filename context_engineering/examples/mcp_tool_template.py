"""
MCP 工具實現範例模板

此文件提供 MCP 工具實現的標準模式和最佳實踐。
所有新的 MCP 工具都應該遵循此模板的結構和風格。
"""

import logging
from datetime import datetime
from typing import Dict, Any, List, Optional, Union

# 假設的 FastMCP 裝飾器
def mcp_tool():
    """MCP 工具裝飾器模擬"""
    def decorator(func):
        return func
    return decorator

# 設置日誌
logger = logging.getLogger(__name__)

# 範例 1: 基本 MCP 工具模板
@mcp_tool()
def basic_mcp_tool(param1: str, param2: int = 10) -> Dict[str, Any]:
    """
    基本 MCP 工具實現範例
    
    Args:
        param1: 必需的字串參數
        param2: 可選的整數參數，預設值為 10
        
    Returns:
        Dict[str, Any]: 標準化的回傳結果
    """
    logger.info(f"basic_mcp_tool called with param1='{param1}', param2={param2}")
    
    try:
        # 1. 參數驗證
        if not param1:
            return {
                "status": "error",
                "message": "param1 is required",
                "error_code": "PARAM_ERROR",
                "suggestion": "Please provide a valid param1 value"
            }
        
        if param2 < 0:
            return {
                "status": "error",
                "message": "param2 must be non-negative",
                "error_code": "PARAM_ERROR",
                "suggestion": "Please provide a non-negative value for param2"
            }
        
        # 2. 核心邏輯實現
        result = perform_basic_operation(param1, param2)
        
        # 3. 成功回傳
        return {
            "status": "success",
            "result": result,
            "message": f"Operation completed successfully with {len(result)} items",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in basic_mcp_tool: {e}")
        return {
            "status": "error",
            "message": str(e),
            "error_code": "EXECUTION_ERROR",
            "suggestion": "Check logs for detailed error information"
        }

def perform_basic_operation(param1: str, param2: int) -> List[Dict[str, Any]]:
    """核心邏輯實現"""
    return [{"name": f"item_{i}", "value": f"{param1}_{i}"} for i in range(param2)]

# 範例 2: AutoCAD 整合 MCP 工具
@mcp_tool()
def autocad_mcp_tool(x: float, y: float, layer: str = None) -> Dict[str, Any]:
    """
    AutoCAD 整合 MCP 工具範例
    
    Args:
        x: X 座標
        y: Y 座標
        layer: 圖層名稱（可選）
        
    Returns:
        Dict[str, Any]: 操作結果
    """
    logger.info(f"autocad_mcp_tool called with x={x}, y={y}, layer={layer}")
    
    try:
        # 1. 參數驗證
        if not isinstance(x, (int, float)) or not isinstance(y, (int, float)):
            return {
                "status": "error",
                "message": "x and y must be numeric values",
                "error_code": "PARAM_ERROR"
            }
        
        # 2. 取得 AutoCAD 工具實例
        autocad_util = get_autocad_util()
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD utility not available",
                "error_code": "AUTOCAD_ERROR",
                "suggestion": "Please ensure AutoCAD is running and connected"
            }
        
        # 3. 檢查 AutoCAD 連接
        if not autocad_util.connected_autocad():
            try:
                # 嘗試連接 AutoCAD
                mock_main_body = MockMainBody()
                autocad_util.connect_autocad(mock_main_body)
                
                if not autocad_util.connected_autocad():
                    return {
                        "status": "error",
                        "message": "Cannot connect to AutoCAD",
                        "error_code": "AUTOCAD_CONNECTION_ERROR",
                        "suggestion": "Please ensure AutoCAD is running"
                    }
            except Exception as e:
                return {
                    "status": "error",
                    "message": f"Failed to connect to AutoCAD: {str(e)}",
                    "error_code": "AUTOCAD_CONNECTION_ERROR"
                }
        
        # 4. 執行 AutoCAD 操作
        result = perform_autocad_operation(autocad_util, x, y, layer)
        
        return {
            "status": "success",
            "result": result,
            "message": "AutoCAD operation completed successfully",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in autocad_mcp_tool: {e}")
        return {
            "status": "error",
            "message": str(e),
            "error_code": "EXECUTION_ERROR"
        }

def perform_autocad_operation(autocad_util, x: float, y: float, layer: str = None) -> Dict[str, Any]:
    """執行 AutoCAD 操作"""
    # 模擬 AutoCAD 操作
    return {
        "operation": "create_point",
        "coordinates": [x, y],
        "layer": layer or "0",
        "entity_id": f"point_{x}_{y}"
    }

# 範例 3: Odoo 整合 MCP 工具
@mcp_tool()
def odoo_mcp_tool(project_id: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Odoo 整合 MCP 工具範例
    
    Args:
        project_id: Odoo 專案 ID
        data: 要同步的資料
        
    Returns:
        Dict[str, Any]: 同步結果
    """
    logger.info(f"odoo_mcp_tool called with project_id={project_id}")
    
    try:
        # 1. 參數驗證
        if not project_id:
            return {
                "status": "error",
                "message": "project_id is required",
                "error_code": "PARAM_ERROR"
            }
        
        if not data:
            return {
                "status": "error",
                "message": "data is required",
                "error_code": "PARAM_ERROR"
            }
        
        # 2. 取得 Odoo 工具實例
        odoo_util = get_odoo_util()
        if not odoo_util:
            return {
                "status": "error",
                "message": "Odoo utility not available",
                "error_code": "ODOO_ERROR",
                "suggestion": "Please check Odoo connection configuration"
            }
        
        # 3. 檢查 Odoo 連接
        if not odoo_util.connected_odoo():
            try:
                odoo_connection = get_odoo_connection_config()
                if not odoo_connection:
                    return {
                        "status": "error",
                        "message": "No Odoo configuration found",
                        "error_code": "ODOO_CONFIG_ERROR"
                    }
                
                odoo_util.connect_odoo(odoo_connection)
                if not odoo_util.connected_odoo():
                    return {
                        "status": "error",
                        "message": "Failed to connect to Odoo",
                        "error_code": "ODOO_CONNECTION_ERROR"
                    }
            except Exception as e:
                return {
                    "status": "error",
                    "message": f"Odoo connection failed: {str(e)}",
                    "error_code": "ODOO_CONNECTION_ERROR"
                }
        
        # 4. 執行 Odoo 操作
        result = perform_odoo_operation(odoo_util, project_id, data)
        
        return {
            "status": "success",
            "result": result,
            "message": "Odoo operation completed successfully",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in odoo_mcp_tool: {e}")
        return {
            "status": "error",
            "message": str(e),
            "error_code": "EXECUTION_ERROR"
        }

def perform_odoo_operation(odoo_util, project_id: str, data: Dict[str, Any]) -> Dict[str, Any]:
    """執行 Odoo 操作"""
    # 模擬 Odoo 操作
    return {
        "operation": "sync_data",
        "project_id": project_id,
        "items_synced": len(data),
        "sync_id": f"sync_{project_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    }

# 輔助類別和函數
class MockMainBody:
    """模擬 GUI 主體組件"""
    def winfo_children(self):
        return []

def get_autocad_util():
    """取得 AutoCAD 工具實例"""
    # 在實際實現中，這會返回真正的 AutoCAD 工具實例
    return MockAutoCADUtil()

def get_odoo_util():
    """取得 Odoo 工具實例"""
    # 在實際實現中，這會返回真正的 Odoo 工具實例
    return MockOdooUtil()

def get_odoo_connection_config():
    """取得 Odoo 連接配置"""
    # 在實際實現中，這會從資料庫讀取真正的配置
    return {"host": "example.com", "db_name": "test", "url": "http://example.com", "token": "test_token"}

class MockAutoCADUtil:
    """模擬 AutoCAD 工具類別"""
    def connected_autocad(self):
        return True
    
    def connect_autocad(self, main_body):
        return True

class MockOdooUtil:
    """模擬 Odoo 工具類別"""
    def connected_odoo(self):
        return True
    
    def connect_odoo(self, connection):
        return True

# 標準化的錯誤處理模式
class MCPError(Exception):
    """MCP 工具專用例外類別"""
    def __init__(self, message: str, error_code: str, suggestion: str = None):
        self.message = message
        self.error_code = error_code
        self.suggestion = suggestion
        super().__init__(message)

def handle_mcp_error(error: MCPError) -> Dict[str, Any]:
    """標準化的 MCP 錯誤處理"""
    response = {
        "status": "error",
        "message": error.message,
        "error_code": error.error_code
    }
    
    if error.suggestion:
        response["suggestion"] = error.suggestion
    
    return response

# 標準化的成功回應
def create_success_response(result: Any, message: str = "Operation completed successfully") -> Dict[str, Any]:
    """創建標準化的成功回應"""
    return {
        "status": "success",
        "result": result,
        "message": message,
        "timestamp": datetime.now().isoformat()
    }

# 參數驗證輔助函數
def validate_required_params(**params) -> Optional[Dict[str, Any]]:
    """驗證必需參數"""
    for name, value in params.items():
        if value is None or value == "":
            return {
                "status": "error",
                "message": f"Parameter '{name}' is required",
                "error_code": "PARAM_ERROR",
                "suggestion": f"Please provide a valid value for '{name}'"
            }
    return None

def validate_numeric_params(**params) -> Optional[Dict[str, Any]]:
    """驗證數值參數"""
    for name, value in params.items():
        if not isinstance(value, (int, float)):
            return {
                "status": "error",
                "message": f"Parameter '{name}' must be numeric",
                "error_code": "PARAM_ERROR",
                "suggestion": f"Please provide a numeric value for '{name}'"
            }
    return None

# 使用範例
if __name__ == "__main__":
    # 測試基本工具
    result1 = basic_mcp_tool("test", 5)
    print("Basic tool result:", result1)
    
    # 測試 AutoCAD 工具
    result2 = autocad_mcp_tool(10.0, 20.0, "Layer1")
    print("AutoCAD tool result:", result2)
    
    # 測試 Odoo 工具
    result3 = odoo_mcp_tool("project_123", {"item1": "value1", "item2": "value2"})
    print("Odoo tool result:", result3)