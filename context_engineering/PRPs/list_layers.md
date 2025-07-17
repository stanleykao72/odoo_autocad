# PRP: list_layers

> **功能**: 列出 AutoCAD 中所有可用的圖層  
> **類型**: MCP 工具  
> **優先級**: 中  
> **日期**: 2025年7月16日  

## 功能概述

### 目標
列出 AutoCAD 中所有可用的圖層及其屬性，支援過濾和排序功能。這是第一階段圖層管理工具的重要組成部分。

### 使用場景
1. **圖層查看**: 工程師需要查看圖面中所有的圖層
2. **圖層管理**: 了解圖層的狀態和屬性
3. **圖層過濾**: 根據條件過濾特定圖層
4. **AI 助手整合**: 透過自然語言查詢圖層資訊

### 成功標準
- [ ] 能夠列出所有圖層及其屬性
- [ ] 支援過濾和排序功能
- [ ] 提供詳細的圖層資訊
- [ ] 支援當前圖層標識
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def list_layers(
    filter_type: str = "all",
    sort_by: str = "name",
    include_details: bool = True
) -> Dict[str, Any]:
    """列出 AutoCAD 中所有可用的圖層"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| filter_type | str | 否 | "all" | 過濾類型 ("all", "visible", "current", "frozen", "locked") |
| sort_by | str | 否 | "name" | 排序方式 ("name", "color", "created") |
| include_details | bool | 否 | True | 是否包含詳細屬性 |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "layers": [
            {
                "name": "0",
                "color": 7,
                "is_current": True,
                "linetype": "Continuous",
                "lineweight": "Default",
                "on": True,
                "frozen": False,
                "locked": False,
                "description": "Default layer"
            },
            {
                "name": "WALLS",
                "color": 2,
                "is_current": False,
                "linetype": "Continuous",
                "lineweight": "Default",
                "on": True,
                "frozen": False,
                "locked": False,
                "description": "Wall elements"
            }
        ],
        "total_count": 2,
        "current_layer": "0",
        "filter_applied": "all",
        "sort_applied": "name",
        "retrieved_at": "2025-07-16T10:30:00"
    },
    "message": "成功列出 2 個圖層",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "無效的過濾類型",
    "error_code": "INVALID_FILTER_TYPE",
    "suggestion": "過濾類型必須是 'all', 'visible', 'current', 'frozen', 'locked' 之一"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `_autocad_util`）
- [ ] 需要 Odoo 連接（僅用於日誌記錄）
- [ ] 需要資料庫存取（僅用於配置）
- [x] 需要 win32com.client 和 pythoncom

### 核心邏輯
1. **參數驗證**: 檢查過濾類型、排序方式
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **圖層掃描**: 獲取所有圖層的基本資訊
4. **屬性收集**: 收集每個圖層的詳細屬性
5. **過濾處理**: 根據過濾條件篩選圖層
6. **排序處理**: 根據排序條件排序圖層列表
7. **結果回傳**: 返回圖層列表和統計資訊

### 錯誤處理
- **參數錯誤**: 過濾類型無效、排序方式無效
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、文檔未開啟
- **掃描錯誤**: 圖層掃描失敗、屬性讀取失敗
- **處理錯誤**: 過濾失敗、排序失敗

## 範例代碼

### 基本實現
```python
@mcp.tool()
def list_layers(
    filter_type: str = "all",
    sort_by: str = "name",
    include_details: bool = True
) -> Dict[str, Any]:
    """列出 AutoCAD 中所有可用的圖層"""
    logger.info(f"list_layers called with params: {locals()}")
    
    try:
        # 參數驗證
        valid_filters = ["all", "visible", "current", "frozen", "locked"]
        if filter_type not in valid_filters:
            return {
                "status": "error",
                "message": "無效的過濾類型",
                "error_code": "INVALID_FILTER_TYPE",
                "suggestion": f"過濾類型必須是 {valid_filters} 之一"
            }
        
        valid_sorts = ["name", "color", "created"]
        if sort_by not in valid_sorts:
            return {
                "status": "error",
                "message": "無效的排序方式",
                "error_code": "INVALID_SORT_TYPE",
                "suggestion": f"排序方式必須是 {valid_sorts} 之一"
            }
        
        # 檢查 AutoCAD 連接
        autocad_util = _autocad_util
        if not autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 列出圖層
        result = autocad_util.list_layers(
            filter_type=filter_type,
            sort_by=sort_by,
            include_details=include_details
        )
        
        return {
            "status": "success",
            "data": {
                "layers": result.get("layers", []),
                "total_count": result.get("total_count", 0),
                "current_layer": result.get("current_layer"),
                "filter_applied": filter_type,
                "sort_applied": sort_by,
                "retrieved_at": datetime.now().isoformat()
            },
            "message": f"成功列出 {result.get('total_count', 0)} 個圖層",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in list_layers: {e}")
        return {
            "status": "error",
            "message": f"列出圖層失敗: {str(e)}",
            "error_code": "LAYER_LISTING_ERROR"
        }
```

### 使用範例
```python
# 基本使用 - 列出所有圖層
result = list_layers()

# 只列出可見圖層
result = list_layers(filter_type="visible")

# 按顏色排序
result = list_layers(sort_by="color")

# 簡化資訊
result = list_layers(include_details=False)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_list_layers.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime

import mcp_server_fastmcp

class TestListLayers:
    def test_list_layers_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "0",
                    "color": 7,
                    "is_current": True,
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "on": True,
                    "frozen": False,
                    "locked": False,
                    "description": "Default layer"
                },
                {
                    "name": "WALLS",
                    "color": 2,
                    "is_current": False,
                    "linetype": "Continuous",
                    "lineweight": "Default",
                    "on": True,
                    "frozen": False,
                    "locked": False,
                    "description": "Wall elements"
                }
            ],
            "total_count": 2,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers()
        
        # Assert
        assert result["status"] == "success"
        assert len(result["data"]["layers"]) == 2
        assert result["data"]["total_count"] == 2
        assert result["data"]["current_layer"] == "0"
        assert result["data"]["filter_applied"] == "all"
        assert result["data"]["sort_applied"] == "name"
        assert "retrieved_at" in result["data"]
        assert "message" in result
        assert "timestamp" in result
        
    def test_list_layers_parameter_validation_invalid_filter(self):
        """參數驗證測試 - 無效過濾類型"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.list_layers(filter_type="invalid")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_FILTER_TYPE"
        
    def test_list_layers_parameter_validation_invalid_sort(self):
        """參數驗證測試 - 無效排序方式"""
        # Arrange
        mock_autocad_util = Mock()
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.list_layers(sort_by="invalid")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_SORT_TYPE"
        
    def test_list_layers_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試"""
        # Arrange
        mcp_server_fastmcp._autocad_util = None
        
        # Act & Assert
        result = mcp_server_fastmcp.list_layers()
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_list_layers_with_filter_visible(self):
        """過濾測試 - 可見圖層"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {
                    "name": "0",
                    "color": 7,
                    "is_current": True,
                    "on": True,
                    "frozen": False,
                    "locked": False
                }
            ],
            "total_count": 1,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(filter_type="visible")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["filter_applied"] == "visible"
        assert result["data"]["total_count"] == 1
        
    def test_list_layers_with_sort_color(self):
        """排序測試 - 按顏色排序"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [
                {"name": "WALLS", "color": 2},
                {"name": "0", "color": 7}
            ],
            "total_count": 2,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(sort_by="color")
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["sort_applied"] == "color"
        
    def test_list_layers_autocad_util_called_correctly(self):
        """驗證 AutoCAD 工具被正確呼叫"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.return_value = {
            "layers": [],
            "total_count": 0,
            "current_layer": "0"
        }
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act
        result = mcp_server_fastmcp.list_layers(
            filter_type="visible",
            sort_by="color",
            include_details=False
        )
        
        # Assert
        assert result["status"] == "success"
        mock_autocad_util.list_layers.assert_called_once_with(
            filter_type="visible",
            sort_by="color",
            include_details=False
        )
        
    def test_list_layers_exception_handling(self):
        """異常處理測試"""
        # Arrange
        mock_autocad_util = Mock()
        mock_autocad_util.list_layers.side_effect = Exception("Layer listing error")
        mcp_server_fastmcp._autocad_util = mock_autocad_util
        
        # Act & Assert
        result = mcp_server_fastmcp.list_layers()
        assert result["status"] == "error"
        assert result["error_code"] == "LAYER_LISTING_ERROR"
        
    def test_list_layers_function_exists(self):
        """驗證函數存在性"""
        # Act & Assert
        assert hasattr(mcp_server_fastmcp, 'list_layers')
        assert callable(mcp_server_fastmcp.list_layers)
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def list_layers(
    filter_type: str = "all",
    sort_by: str = "name",
    include_details: bool = True
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    # 基本驗證
    valid_filters = ["all", "visible", "current", "frozen", "locked"]
    if filter_type not in valid_filters:
        return {
            "status": "error",
            "message": "無效的過濾類型",
            "error_code": "INVALID_FILTER_TYPE"
        }
    
    valid_sorts = ["name", "color", "created"]
    if sort_by not in valid_sorts:
        return {
            "status": "error",
            "message": "無效的排序方式",
            "error_code": "INVALID_SORT_TYPE"
        }
    
    if not _autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    # 模擬結果
    return {
        "status": "success",
        "data": {
            "layers": [],
            "total_count": 0,
            "current_layer": "0",
            "filter_applied": filter_type,
            "sort_applied": sort_by,
            "retrieved_at": datetime.now().isoformat()
        },
        "message": f"成功列出 0 個圖層",
        "timestamp": datetime.now().isoformat()
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，包含實際 AutoCAD 操作
@mcp.tool()
def list_layers(
    filter_type: str = "all",
    sort_by: str = "name",
    include_details: bool = True
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_list_layers.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_list_layers.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_list_layers.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_list_layers.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_list_layers_integration.py
import pytest
from mcp_server_fastmcp import list_layers

class TestListLayersIntegration:
    def test_list_layers_with_real_autocad(self):
        """與真實 AutoCAD 的整合測試"""
        # 注意：這個測試需要真實的 AutoCAD 連接
        # 假設 AutoCAD 已經啟動並連接
        pass
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
- [ ] 支援所有指定的參數組合
- [ ] 與 AutoCAD 正常互動
- [ ] 過濾和排序功能正確

### 品質標準
- [ ] 代碼覆蓋率 ≥ 95%
- [ ] 所有測試通過
- [ ] 符合代碼風格指南
- [ ] 包含完整日誌記錄
- [ ] 遵循 TDD 原則
- [ ] 無靜態分析警告

### 整合標準
- [ ] 與現有 MCP 系統整合
- [ ] 與 AutoCAD 正常連接
- [ ] 與 Odoo 系統相容
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
- [AutoCAD COM 範例](../examples/autocad_com_examples.py)
- [錯誤處理模式](../examples/error_handling_patterns.py)

### 現有代碼
- `mcp_server_fastmcp.py` - 主要實施位置
- `utility/util_autocad.py` - AutoCAD 工具類別
- `tests/unit/test_mcp_*.py` - 現有測試參考

---

**注意**: 此 PRP 應該作為 AI 助手的完整上下文使用，確保生成的代碼符合所有要求和標準。實施時請嚴格遵循 TDD 原則：先寫測試，再寫實現。