# PRP: create_new_drawing

> **功能**: 創建新的 AutoCAD 圖面檔案  
> **類型**: MCP 工具  
> **優先級**: 高  
> **日期**: 2025年7月16日  

## 功能概述

### 目標
創建新的 AutoCAD 圖面檔案，支援自定義模板和單位設定，並整合現有的 AutoCAD 連接邏輯。這是第一階段基礎繪圖工具的核心功能。

### 使用場景
1. **新專案開始**: 使用者需要建立新的工程圖面
2. **模板應用**: 基於現有模板創建標準化圖面
3. **單位設定**: 根據專案需求設定公制或英制單位
4. **AI 助手整合**: 透過自然語言指令創建圖面

### 成功標準
- [ ] 能夠成功創建新的 AutoCAD 圖面檔案
- [ ] 支援自定義模板路徑
- [ ] 正確設定圖面單位系統
- [ ] 提供清晰的操作回饋
- [ ] 與現有 AutoCAD 連接邏輯完全整合
- [ ] 通過所有 TDD 測試（95%+ 覆蓋率）

## 技術規格

### 函數簽名
```python
@mcp.tool()
def create_new_drawing(
    drawing_name: str = "新圖面",
    template_path: str = "",
    units: str = "公制",
    save_path: str = ""
) -> Dict[str, Any]:
    """創建新的 AutoCAD 圖面檔案"""
```

### 參數規格
| 參數名 | 類型 | 必需 | 預設值 | 描述 |
|--------|------|------|--------|------|
| drawing_name | str | 否 | "新圖面" | 圖面名稱 |
| template_path | str | 否 | "" | 模板檔案路徑（空字串使用預設模板） |
| units | str | 否 | "公制" | 單位系統（公制/英制） |
| save_path | str | 否 | "" | 儲存路徑（空字串使用預設路徑） |

### 回傳值規格
```python
# 成功回傳
{
    "status": "success",
    "data": {
        "drawing_name": "新圖面.dwg",
        "full_path": "C:\\Projects\\新圖面.dwg",
        "template_used": "預設模板",
        "units": "公制",
        "created_at": "2025-07-16T10:30:00"
    },
    "message": "成功創建圖面: 新圖面.dwg",
    "timestamp": "2025-07-16T10:30:00"
}

# 錯誤回傳
{
    "status": "error",
    "message": "AutoCAD 連接未建立",
    "error_code": "AUTOCAD_NOT_CONNECTED",
    "suggestion": "請確保 AutoCAD 正在運行並已建立連接"
}
```

## 實施要求

### 依賴項目
- [x] 需要 AutoCAD 連接（使用現有的 `self.autocad_util`）
- [ ] 需要 Odoo 連接（僅用於日誌記錄）
- [ ] 需要資料庫存取（僅用於配置）
- [x] 需要 win32com.client 和 pythoncom

### 核心邏輯
1. **參數驗證**: 檢查圖面名稱格式、模板路徑有效性、單位系統合法性
2. **AutoCAD 連接檢查**: 確保 AutoCAD 應用程式正在運行且可用
3. **模板處理**: 如果提供模板路徑，驗證模板存在性；否則使用預設模板
4. **圖面創建**: 使用 AutoCAD COM API 創建新圖面
5. **單位設定**: 根據參數設定圖面單位系統
6. **儲存處理**: 如果提供儲存路徑，將圖面儲存到指定位置
7. **結果回傳**: 返回創建的圖面資訊

### 錯誤處理
- **參數錯誤**: 無效的圖面名稱、不存在的模板路徑、不支援的單位系統
- **連接錯誤**: AutoCAD 未運行、COM 連接失敗、權限問題
- **檔案錯誤**: 模板檔案損壞、儲存路徑無效、磁碟空間不足
- **執行錯誤**: AutoCAD 內部錯誤、COM 操作失敗

## 範例代碼

### 基本實現
```python
@mcp.tool()
def create_new_drawing(
    drawing_name: str = "新圖面",
    template_path: str = "",
    units: str = "公制",
    save_path: str = ""
) -> Dict[str, Any]:
    """創建新的 AutoCAD 圖面檔案"""
    logger.info(f"create_new_drawing called with params: {locals()}")
    
    try:
        # 參數驗證
        if not drawing_name.strip():
            return {
                "status": "error",
                "message": "圖面名稱不能為空",
                "error_code": "INVALID_DRAWING_NAME"
            }
        
        if units not in ["公制", "英制"]:
            return {
                "status": "error",
                "message": "單位系統必須是 '公制' 或 '英制'",
                "error_code": "INVALID_UNITS"
            }
        
        # 檢查 AutoCAD 連接
        if not self.autocad_util:
            return {
                "status": "error",
                "message": "AutoCAD 連接未建立",
                "error_code": "AUTOCAD_NOT_CONNECTED"
            }
        
        # 驗證模板檔案
        if template_path and not os.path.exists(template_path):
            return {
                "status": "error",
                "message": f"模板檔案不存在: {template_path}",
                "error_code": "TEMPLATE_NOT_FOUND"
            }
        
        # 創建新圖面
        result = self.autocad_util.create_new_drawing(
            drawing_name=drawing_name,
            template_path=template_path or None,
            units=units,
            save_path=save_path or None
        )
        
        return {
            "status": "success",
            "data": {
                "drawing_name": result.get("drawing_name"),
                "full_path": result.get("full_path"),
                "template_used": template_path or "預設模板",
                "units": units,
                "created_at": datetime.now().isoformat()
            },
            "message": f"成功創建圖面: {result.get('drawing_name')}",
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Error in create_new_drawing: {e}")
        return {
            "status": "error",
            "message": f"創建圖面失敗: {str(e)}",
            "error_code": "DRAWING_CREATION_ERROR"
        }
```

### 使用範例
```python
# 基本使用 - 創建預設圖面
result = create_new_drawing()

# 使用自定義名稱和單位
result = create_new_drawing(
    drawing_name="工程圖_A1",
    units="公制"
)

# 使用模板創建圖面
result = create_new_drawing(
    drawing_name="標準圖面",
    template_path="C:\\Templates\\standard.dwt",
    units="公制",
    save_path="C:\\Projects\\標準圖面.dwg"
)
```

## TDD 測試要求

### Red-Green-Refactor 循環

#### 1. Red 階段 (失敗測試)
```python
# 先寫失敗測試 - tests/unit/test_create_new_drawing.py
import pytest
from unittest.mock import patch, Mock
from datetime import datetime
from mcp_server_fastmcp import FastMCPServer

class TestCreateNewDrawing:
    def test_create_new_drawing_basic_functionality(self):
        """基本功能測試 - 應該失敗"""
        # Arrange
        server = FastMCPServer()
        server.autocad_util = Mock()
        server.autocad_util.create_new_drawing.return_value = {
            "drawing_name": "新圖面.dwg",
            "full_path": "C:\\Projects\\新圖面.dwg"
        }
        
        expected = {
            "status": "success",
            "data": {
                "drawing_name": "新圖面.dwg",
                "full_path": "C:\\Projects\\新圖面.dwg",
                "template_used": "預設模板",
                "units": "公制"
            }
        }
        
        # Act
        result = server.create_new_drawing()
        
        # Assert
        assert result["status"] == "success"
        assert result["data"]["drawing_name"] == "新圖面.dwg"
        assert result["data"]["units"] == "公制"
        
    def test_create_new_drawing_parameter_validation(self):
        """參數驗證測試 - 應該失敗"""
        server = FastMCPServer()
        server.autocad_util = Mock()
        
        # 測試空白圖面名稱
        result = server.create_new_drawing(drawing_name="")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_DRAWING_NAME"
        
        # 測試無效單位
        result = server.create_new_drawing(units="無效單位")
        assert result["status"] == "error"
        assert result["error_code"] == "INVALID_UNITS"
        
    def test_create_new_drawing_autocad_connection_error(self):
        """AutoCAD 連接錯誤測試 - 應該失敗"""
        server = FastMCPServer()
        server.autocad_util = None
        
        result = server.create_new_drawing()
        assert result["status"] == "error"
        assert result["error_code"] == "AUTOCAD_NOT_CONNECTED"
        
    def test_create_new_drawing_template_not_found(self):
        """模板檔案不存在測試 - 應該失敗"""
        server = FastMCPServer()
        server.autocad_util = Mock()
        
        with patch('os.path.exists', return_value=False):
            result = server.create_new_drawing(template_path="non_existent.dwt")
            assert result["status"] == "error"
            assert result["error_code"] == "TEMPLATE_NOT_FOUND"
            
    def test_create_new_drawing_autocad_util_exception(self):
        """AutoCAD 工具異常測試 - 應該失敗"""
        server = FastMCPServer()
        server.autocad_util = Mock()
        server.autocad_util.create_new_drawing.side_effect = Exception("AutoCAD error")
        
        result = server.create_new_drawing()
        assert result["status"] == "error"
        assert result["error_code"] == "DRAWING_CREATION_ERROR"
```

#### 2. Green 階段 (最小實現)
```python
# 最小實現讓測試通過
@mcp.tool()
def create_new_drawing(
    drawing_name: str = "新圖面",
    template_path: str = "",
    units: str = "公制",
    save_path: str = ""
) -> Dict[str, Any]:
    """最小實現 - 僅讓測試通過"""
    if not drawing_name.strip():
        return {
            "status": "error",
            "message": "圖面名稱不能為空",
            "error_code": "INVALID_DRAWING_NAME"
        }
    
    if units not in ["公制", "英制"]:
        return {
            "status": "error",
            "message": "單位系統必須是 '公制' 或 '英制'",
            "error_code": "INVALID_UNITS"
        }
    
    if not self.autocad_util:
        return {
            "status": "error",
            "message": "AutoCAD 連接未建立",
            "error_code": "AUTOCAD_NOT_CONNECTED"
        }
    
    return {
        "status": "success",
        "data": {
            "drawing_name": "新圖面.dwg",
            "full_path": "C:\\Projects\\新圖面.dwg",
            "template_used": "預設模板",
            "units": "公制"
        }
    }
```

#### 3. Refactor 階段 (完整實現)
```python
# 完整實現，保持測試通過
@mcp.tool()
def create_new_drawing(
    drawing_name: str = "新圖面",
    template_path: str = "",
    units: str = "公制",
    save_path: str = ""
) -> Dict[str, Any]:
    """完整實現 - 包含所有邏輯"""
    # 完整的實現邏輯（如上方基本實現）
```

### TDD 測試指令
```bash
# 1. 寫失敗測試
python -m pytest tests/unit/test_create_new_drawing.py -v  # 應該失敗

# 2. 最小實現
# 修改 mcp_server_fastmcp.py 實現最小代碼

# 3. 測試通過
python -m pytest tests/unit/test_create_new_drawing.py -v  # 應該通過

# 4. 重構並保持測試通過
python -m pytest tests/unit/test_create_new_drawing.py -v  # 持續通過

# 5. 覆蓋率檢查
python -m pytest tests/unit/test_create_new_drawing.py --cov=mcp_server_fastmcp --cov-report=term-missing
```

### 整合測試 (TDD 方式)
```python
# tests/integration/test_create_new_drawing_integration.py
import pytest
from mcp_server_fastmcp import FastMCPServer

class TestCreateNewDrawingIntegration:
    def test_create_new_drawing_with_real_autocad(self):
        """與真實 AutoCAD 的整合測試"""
        # 注意：這個測試需要真實的 AutoCAD 連接
        server = FastMCPServer()
        # 假設 AutoCAD 已經啟動並連接
        if server.autocad_util:
            result = server.create_new_drawing(
                drawing_name="測試圖面",
                units="公制"
            )
            assert result["status"] == "success"
            assert "測試圖面" in result["data"]["drawing_name"]
        else:
            pytest.skip("AutoCAD 連接不可用")
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