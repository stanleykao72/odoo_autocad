# PRP: sync_drawing_to_odoo

> **功能**: 將 AutoCAD 圖面資訊同步到 Odoo 專案  
> **類型**: MCP 工具  
> **優先級**: 高  
> **日期**: 2025年7月16日  
> **階段**: 第三階段 Odoo 深度整合

## 功能概述

### 目標
將 AutoCAD 圖面資訊完整同步到 Odoo 專案系統中，包括圖面元素、參數資料、BOQ 資訊等。支援自動建立專案關聯，並提供雙向同步功能，確保 AutoCAD 和 Odoo 系統的資料一致性。

### 使用場景
1. **專案建立**: 從 AutoCAD 圖面自動建立 Odoo 專案
2. **資料同步**: 將圖面變更同步到 Odoo 系統
3. **進度追蹤**: 同步專案進度和狀態
4. **團隊協作**: 多人同時操作時的資料同步
5. **AI 助手整合**: 透過自然語言指令進行圖面同步

### 成功標準
- [ ] 能夠從 AutoCAD 圖面建立 Odoo 專案
- [ ] 支援圖面元素資訊同步
- [ ] 支援 BOQ 資料同步
- [ ] 支援雙向同步功能
- [ ] 提供詳細的同步報告
- [ ] 支援增量同步機制
- [ ] 與現有 Odoo 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def sync_drawing_to_odoo(
    drawing_name: str,
    project_name: str = None,
    project_id: int = None,
    sync_mode: str = "full",
    create_project: bool = True,
    sync_elements: bool = True,
    sync_boq: bool = True,
    sync_parameters: bool = True
) -> Dict[str, Any]:
    """將 AutoCAD 圖面資訊同步到 Odoo 專案"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| drawing_name | str | 是 | - | 圖面名稱 |
| project_name | str | 否 | None | Odoo 專案名稱 |
| project_id | int | 否 | None | Odoo 專案 ID |
| sync_mode | str | 否 | "full" | 同步模式 ("full", "incremental", "elements_only") |
| create_project | bool | 否 | True | 如果專案不存在是否創建 |
| sync_elements | bool | 否 | True | 是否同步圖面元素 |
| sync_boq | bool | 否 | True | 是否同步 BOQ 資料 |
| sync_parameters | bool | 否 | True | 是否同步參數資料 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "project_info": {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        },
        "sync_summary": {
            "total_elements": 25,
            "synced_elements": 23,
            "failed_elements": 2,
            "sync_mode": "full",
            "sync_time": 2.5
        },
        "sync_details": {
            "elements_sync": {
                "status": "success",
                "synced_count": 23,
                "failed_count": 2,
                "element_breakdown": {
                    "line": 10,
                    "circle": 5,
                    "text": 6,
                    "dimension": 2
                }
            },
            "boq_sync": {
                "status": "success",
                "synced_items": 15,
                "total_amount": 125000.0,
                "currency": "TWD"
            },
            "parameters_sync": {
                "status": "success",
                "synced_params": 8,
                "validation_errors": 0
            }
        },
        "odoo_urls": {
            "project_url": "https://odoo.example.com/project/123",
            "boq_url": "https://odoo.example.com/boq/456"
        },
        "sync_log": [
            "開始同步圖面: 測試圖面",
            "創建新專案: 測試專案",
            "同步元素: 23/25 成功",
            "同步BOQ: 15項完成",
            "同步參數: 8項完成",
            "同步完成"
        ],
        "synced_at": "2025-07-16T10:30:00"
    },
    "message": "圖面同步完成: 23/25 元素成功同步到專案 '測試專案'",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "圖面名稱不能為空",
    "error_code": "INVALID_DRAWING_NAME",
    "suggestion": "請提供有效的圖面名稱"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [x] 需要 Odoo 連接（使用現有的 `_odoo_util`）
- [x] 需要現有的 scan_elements 功能
- [x] 需要現有的 generate_boq 功能
- [ ] 需要 Odoo 專案管理 API
- [ ] 需要資料同步機制

### 核心邏輯
1. **參數驗證**: 檢查圖面名稱、專案資訊、同步模式
2. **系統連接檢查**: 確保 AutoCAD 和 Odoo 都已連接
3. **圖面資料獲取**: 使用現有工具獲取圖面資訊
4. **專案處理**: 建立或更新 Odoo 專案
5. **資料同步**: 按照同步模式進行資料同步
6. **結果驗證**: 驗證同步結果和資料完整性
7. **報告生成**: 生成詳細的同步報告

### 錯誤處理
- **參數錯誤**: 圖面名稱無效、專案參數衝突、同步模式無效
- **連接錯誤**: AutoCAD 或 Odoo 連接失敗
- **資料錯誤**: 圖面資料不完整、格式錯誤
- **同步錯誤**: 網路問題、權限問題、資料衝突

## 範例代碼

### 基本實現
```python
@mcp.tool()
def sync_drawing_to_odoo(
    drawing_name: str,
    project_name: str = None,
    project_id: int = None,
    sync_mode: str = "full",
    create_project: bool = True,
    sync_elements: bool = True,
    sync_boq: bool = True,
    sync_parameters: bool = True
) -> Dict[str, Any]:
    """將 AutoCAD 圖面資訊同步到 Odoo 專案"""
    logger.info(f"sync_drawing_to_odoo called with params: {locals()}")
    
    try:
        # 參數驗證
        if not drawing_name or not drawing_name.strip():
            return {
                "status": "error",
                "message": "圖面名稱不能為空",
                "error_code": "INVALID_DRAWING_NAME",
                "suggestion": "請提供有效的圖面名稱"
            }
        
        drawing_name = drawing_name.strip()
        
        # 驗證同步模式
        valid_modes = ["full", "incremental", "elements_only"]
        if sync_mode not in valid_modes:
            return {
                "status": "error",
                "message": f"無效的同步模式: {sync_mode}",
                "error_code": "INVALID_SYNC_MODE",
                "suggestion": f"同步模式必須是 {valid_modes} 之一"
            }
        
        # 專案參數驗證
        if project_name and project_id:
            return {
                "status": "error",
                "message": "不能同時指定專案名稱和專案ID",
                "error_code": "CONFLICTING_PROJECT_PARAMS",
                "suggestion": "請只指定專案名稱或專案ID其中一個"
            }
        
        if project_name:
            project_name = project_name.strip()
            if not project_name:
                return {
                    "status": "error",
                    "message": "專案名稱不能為空",
                    "error_code": "INVALID_PROJECT_NAME"
                }
        
        if project_id is not None:
            try:
                project_id = int(project_id)
                if project_id <= 0:
                    return {
                        "status": "error",
                        "message": "專案ID必須是正整數",
                        "error_code": "INVALID_PROJECT_ID"
                    }
            except (ValueError, TypeError):
                return {
                    "status": "error",
                    "message": "專案ID必須是整數",
                    "error_code": "INVALID_PROJECT_ID_TYPE"
                }
        
        # 檢查系統連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        odoo_util = _odoo_util
        if not odoo_util:
            return {
                "status": "error",
                "message": "Odoo 連接未建立",
                "error_code": "ODOO_NOT_CONNECTED"
            }
        
        # 開始同步處理
        sync_start_time = datetime.now()
        sync_log = []
        sync_log.append(f"開始同步圖面: {drawing_name}")
        
        # 使用現有工具獲取圖面資訊
        if sync_elements:
            scan_result = scan_elements(
                element_type="all",
                include_geometry=True,
                include_properties=True
            )
            if scan_result["status"] != "success":
                return {
                    "status": "error",
                    "message": f"無法獲取圖面元素: {scan_result.get('message', '')}",
                    "error_code": "SCAN_ELEMENTS_FAILED"
                }
            elements_data = scan_result["data"]
        else:
            elements_data = {"elements": [], "summary": {"total_count": 0}}
        
        # 處理專案
        project_info = odoo_util.sync_drawing_project(
            drawing_name=drawing_name,
            project_name=project_name,
            project_id=project_id,
            create_project=create_project
        )
        
        if project_info.get("created_new"):
            sync_log.append(f"創建新專案: {project_info['project_name']}")
        else:
            sync_log.append(f"使用現有專案: {project_info['project_name']}")
        
        # 同步元素
        elements_sync_result = {"status": "success", "synced_count": 0, "failed_count": 0}
        if sync_elements and elements_data["elements"]:
            elements_sync_result = odoo_util.sync_elements_to_project(
                project_id=project_info["project_id"],
                elements=elements_data["elements"],
                sync_mode=sync_mode
            )
            sync_log.append(f"同步元素: {elements_sync_result['synced_count']}/{len(elements_data['elements'])} 成功")
        
        # 同步BOQ
        boq_sync_result = {"status": "success", "synced_items": 0, "total_amount": 0.0}
        if sync_boq:
            boq_result = generate_boq_from_drawing(
                drawing_name=drawing_name,
                project_id=project_info["project_id"],
                include_autocad_data=True
            )
            if boq_result["status"] == "success":
                boq_sync_result = {
                    "status": "success",
                    "synced_items": boq_result["data"]["boq_summary"]["total_items"],
                    "total_amount": boq_result["data"]["boq_summary"]["total_amount"],
                    "currency": boq_result["data"]["boq_summary"]["currency"]
                }
                sync_log.append(f"同步BOQ: {boq_sync_result['synced_items']}項完成")
            else:
                boq_sync_result = {"status": "failed", "error": boq_result.get("message", "")}
        
        # 同步參數
        params_sync_result = {"status": "success", "synced_params": 0, "validation_errors": 0}
        if sync_parameters:
            params_result = odoo_util.sync_drawing_parameters(
                project_id=project_info["project_id"],
                drawing_name=drawing_name
            )
            params_sync_result = {
                "status": "success",
                "synced_params": params_result.get("synced_count", 0),
                "validation_errors": params_result.get("error_count", 0)
            }
            sync_log.append(f"同步參數: {params_sync_result['synced_params']}項完成")
        
        sync_log.append("同步完成")
        
        # 計算同步時間
        sync_time = (datetime.now() - sync_start_time).total_seconds()
        
        return {
            "status": "success",
            "data": {
                "project_info": project_info,
                "sync_summary": {
                    "total_elements": len(elements_data["elements"]),
                    "synced_elements": elements_sync_result["synced_count"],
                    "failed_elements": elements_sync_result["failed_count"],
                    "sync_mode": sync_mode,
                    "sync_time": sync_time
                },
                "sync_details": {
                    "elements_sync": elements_sync_result,
                    "boq_sync": boq_sync_result,
                    "parameters_sync": params_sync_result
                },
                "odoo_urls": {
                    "project_url": f"{odoo_util.base_url}/project/{project_info['project_id']}",
                    "boq_url": f"{odoo_util.base_url}/boq/{project_info['project_id']}"
                },
                "sync_log": sync_log,
                "synced_at": datetime.now().isoformat()
            },
            "message": f"圖面同步完成: {elements_sync_result['synced_count']}/{len(elements_data['elements'])} 元素成功同步到專案 '{project_info['project_name']}'",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in sync_drawing_to_odoo: {e}")
        return {
            "status": "error",
            "message": f"圖面同步失敗: {str(e)}",
            "error_code": "DRAWING_SYNC_ERROR"
        }
```

### Odoo 工具類別方法
```python
def sync_drawing_project(self, drawing_name, project_name, project_id, create_project):
    """同步圖面專案"""
    try:
        # 查找或創建專案
        if project_id:
            # 使用現有專案ID
            project = self.get_project_by_id(project_id)
            if not project:
                raise Exception(f"找不到專案 ID: {project_id}")
            return {
                "project_id": project_id,
                "project_name": project["name"],
                "created_new": False,
                "updated_at": datetime.now().isoformat()
            }
        
        elif project_name:
            # 使用專案名稱
            project = self.get_project_by_name(project_name)
            if project:
                return {
                    "project_id": project["id"],
                    "project_name": project["name"],
                    "created_new": False,
                    "updated_at": datetime.now().isoformat()
                }
            elif create_project:
                # 創建新專案
                new_project = self.create_project({
                    "name": project_name,
                    "description": f"從 AutoCAD 圖面 '{drawing_name}' 自動建立",
                    "drawing_name": drawing_name,
                    "state": "draft"
                })
                return {
                    "project_id": new_project["id"],
                    "project_name": new_project["name"],
                    "created_new": True,
                    "updated_at": datetime.now().isoformat()
                }
            else:
                raise Exception(f"找不到專案 '{project_name}' 且不允許創建新專案")
        
        else:
            # 使用圖面名稱作為專案名稱
            project_name = drawing_name
            project = self.get_project_by_name(project_name)
            if project:
                return {
                    "project_id": project["id"],
                    "project_name": project["name"],
                    "created_new": False,
                    "updated_at": datetime.now().isoformat()
                }
            elif create_project:
                new_project = self.create_project({
                    "name": project_name,
                    "description": f"從 AutoCAD 圖面 '{drawing_name}' 自動建立",
                    "drawing_name": drawing_name,
                    "state": "draft"
                })
                return {
                    "project_id": new_project["id"],
                    "project_name": new_project["name"],
                    "created_new": True,
                    "updated_at": datetime.now().isoformat()
                }
            else:
                raise Exception("未指定專案且不允許創建新專案")
                
    except Exception as e:
        self.log.safe_log_insert(f"同步圖面專案時發生錯誤: {str(e)}\n")
        raise e

def sync_elements_to_project(self, project_id, elements, sync_mode):
    """同步元素到專案"""
    try:
        synced_count = 0
        failed_count = 0
        element_breakdown = {}
        
        for element in elements:
            try:
                # 將元素同步到 Odoo
                element_data = {
                    "project_id": project_id,
                    "element_type": element.get("type", "unknown"),
                    "element_id": element.get("id", ""),
                    "layer": element.get("layer", "0"),
                    "properties": element.get("properties", {}),
                    "geometry": element.get("geometry", {}),
                    "sync_mode": sync_mode
                }
                
                if sync_mode == "incremental":
                    # 檢查元素是否已存在
                    existing = self.get_project_element(project_id, element["id"])
                    if existing:
                        self.update_project_element(existing["id"], element_data)
                    else:
                        self.create_project_element(element_data)
                else:
                    # 完整同步
                    self.create_project_element(element_data)
                
                synced_count += 1
                element_type = element.get("type", "unknown")
                element_breakdown[element_type] = element_breakdown.get(element_type, 0) + 1
                
            except Exception as e:
                failed_count += 1
                self.log.safe_log_insert(f"同步元素失敗: {element.get('id', 'unknown')} - {str(e)}\n")
        
        self.log.safe_log_insert(f"元素同步完成: {synced_count} 成功, {failed_count} 失敗\n")
        
        return {
            "status": "success",
            "synced_count": synced_count,
            "failed_count": failed_count,
            "element_breakdown": element_breakdown
        }
        
    except Exception as e:
        self.log.safe_log_insert(f"同步元素到專案時發生錯誤: {str(e)}\n")
        raise e

def sync_drawing_parameters(self, project_id, drawing_name):
    """同步圖面參數"""
    try:
        # 從 AutoCAD 獲取參數
        params_result = extract_autocad_parameters(
            drawing_path=None,
            use_current_drawing=True
        )
        
        if params_result["status"] != "success":
            raise Exception(f"無法獲取圖面參數: {params_result.get('message', '')}")
        
        parameters = params_result["data"]["parameters"]
        synced_count = 0
        error_count = 0
        
        for param in parameters:
            try:
                param_data = {
                    "project_id": project_id,
                    "drawing_name": drawing_name,
                    "parameter_name": param["name"],
                    "parameter_value": param["value"],
                    "parameter_type": param["type"],
                    "unit": param.get("unit", ""),
                    "description": param.get("description", "")
                }
                
                # 檢查參數是否已存在
                existing = self.get_project_parameter(project_id, param["name"])
                if existing:
                    self.update_project_parameter(existing["id"], param_data)
                else:
                    self.create_project_parameter(param_data)
                
                synced_count += 1
                
            except Exception as e:
                error_count += 1
                self.log.safe_log_insert(f"同步參數失敗: {param['name']} - {str(e)}\n")
        
        self.log.safe_log_insert(f"參數同步完成: {synced_count} 成功, {error_count} 失敗\n")
        
        return {
            "synced_count": synced_count,
            "error_count": error_count
        }
        
    except Exception as e:
        self.log.safe_log_insert(f"同步圖面參數時發生錯誤: {str(e)}\n")
        raise e
```

### 使用範例
```python
# 基本同步
result = sync_drawing_to_odoo(
    drawing_name="建築平面圖",
    project_name="新建專案"
)

# 更新現有專案
result = sync_drawing_to_odoo(
    drawing_name="建築平面圖",
    project_id=123,
    sync_mode="incremental"
)

# 僅同步元素
result = sync_drawing_to_odoo(
    drawing_name="詳細圖",
    project_name="現有專案",
    sync_mode="elements_only",
    sync_boq=False,
    sync_parameters=False
)

# 完整同步
result = sync_drawing_to_odoo(
    drawing_name="總圖",
    project_name="大型專案",
    sync_mode="full",
    create_project=True,
    sync_elements=True,
    sync_boq=True,
    sync_parameters=True
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_sync_drawing_to_odoo.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime

import mcp_server_fastmcp

class TestSyncDrawingToOdoo:
    def test_sync_drawing_to_odoo_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_odoo_util = Mock()
        mock_odoo_util.sync_drawing_project.return_value = {
            "project_id": 123,
            "project_name": "測試專案",
            "created_new": True,
            "updated_at": "2025-07-16T10:30:00"
        }
        mock_odoo_util.sync_elements_to_project.return_value = {
            "status": "success",
            "synced_count": 23,
            "failed_count": 2,
            "element_breakdown": {"line": 10, "circle": 5, "text": 6, "dimension": 2}
        }
        mock_odoo_util.base_url = "https://odoo.example.com"
        
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        mcp_server_fastmcp._odoo_util = mock_odoo_util
        
        # Act
        result = mcp_server_fastmcp.sync_drawing_to_odoo(
            drawing_name="測試圖面",
            project_name="測試專案"
        )
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["project_info"]["project_id"] == 123
        assert result["data"]["project_info"]["project_name"] == "測試專案"
        assert result["data"]["project_info"]["created_new"] == True
        assert result["data"]["sync_summary"]["synced_elements"] == 23
        assert result["data"]["sync_summary"]["failed_elements"] == 2
        assert "sync_details" in result["data"]
        assert "odoo_urls" in result["data"]
        assert "sync_log" in result["data"]
        assert "synced_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def sync_drawing_to_odoo(
    drawing_name: str,
    project_name: str = None,
    project_id: int = None,
    sync_mode: str = "full",
    create_project: bool = True,
    sync_elements: bool = True,
    sync_boq: bool = True,
    sync_parameters: bool = True
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    if not drawing_name or not drawing_name.strip():
        return {
            "status": "error",
            "message": "圖面名稱不能為空",
            "error_code": "INVALID_DRAWING_NAME"
        }
    
    if not _autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    if not _odoo_util:
        return {
            "status": "error",
            "message": "Odoo 連接未建立",
            "error_code": "ODOO_NOT_CONNECTED"
        }
    
    # 模擬結果
    return {
        "status": "success",
        "data": {
            "project_info": {
                "project_id": 123,
                "project_name": project_name or drawing_name,
                "created_new": True,
                "updated_at": datetime.now().isoformat()
            },
            "sync_summary": {
                "total_elements": 25,
                "synced_elements": 23,
                "failed_elements": 2,
                "sync_mode": sync_mode,
                "sync_time": 2.5
            },
            "sync_details": {
                "elements_sync": {"status": "success", "synced_count": 23, "failed_count": 2},
                "boq_sync": {"status": "success", "synced_items": 15, "total_amount": 125000.0},
                "parameters_sync": {"status": "success", "synced_params": 8, "validation_errors": 0}
            },
            "odoo_urls": {
                "project_url": "https://odoo.example.com/project/123",
                "boq_url": "https://odoo.example.com/boq/456"
            },
            "sync_log": ["開始同步圖面: " + drawing_name, "同步完成"],
            "synced_at": datetime.now().isoformat()
        },
        "message": f"圖面同步完成: 23/25 元素成功同步到專案 '{project_name or drawing_name}'",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 Odoo 操作
@mcp.tool()
def sync_drawing_to_odoo(
    drawing_name: str,
    project_name: str = None,
    project_id: int = None,
    sync_mode: str = "full",
    create_project: bool = True,
    sync_elements: bool = True,
    sync_boq: bool = True,
    sync_parameters: bool = True
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_sync_drawing_to_odoo.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_sync_drawing_to_odoo.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_sync_drawing_to_odoo.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_sync_drawing_to_odoo.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 測試覆蓋率要求
- 單元測試覆蓋率: 95%+
- 整合測試覆蓋率: 90%+
- 所有錯誤路徑都要測試
- 所有參數組合都要測試

## 驗證標準

### 功能驗證
- [ ] 基本功能正常運作
- [ ] 參數驗證正確
- [ ] 錯誤處理完整
- [ ] 回傳值格式正確
- [ ] 支援所有指定的同步模式
- [ ] 與 AutoCAD 和 Odoo 正常互動
- [ ] 專案創建和更新正確
- [ ] 資料同步準確
- [ ] 支援雙向同步
- [ ] 同步報告完整

### 品質標準
- [ ] 代碼覆蓋率 ≥ 95%
- [ ] 所有測試通過
- [ ] 符合代碼風格指南
- [ ] 包含完整日誌記錄
- [ ] 遵循 TDD 原則
- [ ] 無靜態分析警告

### 整合標準
- [ ] 與現有 MCP 系統整合
- [ ] 與 AutoCAD 和 Odoo 正常連接
- [ ] 與其他 MCP 工具相容
- [ ] 符合 CLAUDE.md 規則
- [ ] 不影響現有功能
- [ ] 日誌格式一致

## 實施檢查清單

### 開發前
- [ ] 完成 PRP 審查
- [ ] 準備測試資料
- [ ] 確認依賴項目
- [ ] 設定 TDD 環境

### 開發中 (TDD 循環)
- [ ] 寫失敗測試 (Red)
- [ ] 最小實現 (Green)
- [ ] 重構改善 (Refactor)
- [ ] 遵循既有模式
- [ ] 實施錯誤處理
- [ ] 添加日誌記錄

### 開發後
- [ ] 執行所有測試
- [ ] 檢查代碼覆蓋率
- [ ] 進行代碼審查
- [ ] 更新文檔
- [ ] 整合測試
- [ ] 性能測試

## 相關資源

### 參考文件
- [AutoCAD MCP 整合方案](../../doc/AutoCAD_MCP_Integration_Plan.md)
- [CLAUDE.md](../../CLAUDE.md)
- [Context Engineering 分析](../../doc/Context_Engineering_Analysis.md)
- [TDD + Context Engineering 整合指南](../TDD_CONTEXT_ENGINEERING_GUIDE.md)

### 相關範例
- [MCP 工具範例](../examples/mcp_tool_template.py)
- [Odoo 整合範例](../examples/odoo_integration_examples.py)
- [圖面同步模式](../examples/drawing_sync_patterns.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_odoo.py` - Odoo 工具類別
- `utility/util_autocad.py` - AutoCAD 工具類別
- `tests/unit/test_sync_*.py` - 相關測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。