# /execute-prp Command

## 目的
根據指定的 PRP (Product Requirements Prompt) 文檔，實際實施 AutoCAD MCP 功能。

## 使用方法
```
/execute-prp [功能名稱]
```

## 參數
- `[功能名稱]`: 要執行的 PRP 功能名稱（例如：create_new_drawing, draw_line, set_layer）

## 執行步驟

### 1. 讀取 PRP 文檔
- 完整閱讀 `context_engineering/PRPs/[功能名稱].md`
- 理解所有技術要求和規格
- 確認實施標準和驗證要求

### 2. 準備開發環境
- 確認現有系統狀態
- 檢查依賴項目
- 準備測試環境

### 3. 實施功能代碼

#### 主要實施檔案
- **主要實施位置**: `mcp_server_fastmcp.py`
- **輔助工具**: 如需要，在 `utility/` 目錄下創建新的輔助模組

#### 實施標準
- 嚴格遵循 PRP 中的函數簽名
- 使用 `@mcp.tool()` 裝飾器
- 實現完整的參數驗證
- 包含詳細的錯誤處理
- 添加完整的日誌記錄

#### 代碼結構
```python
@mcp.tool()
def [功能名稱]([參數列表]) -> Dict[str, Any]:
    """PRP 中定義的功能描述"""
    logger.info(f"[功能名稱] called with params: {locals()}")
    
    try:
        # 1. 參數驗證
        # 2. 系統連接檢查
        # 3. 核心邏輯實現
        # 4. 結果回傳
        
    except Exception as e:
        logger.error(f"Error in [功能名稱]: {e}")
        return {
            "status": "error",
            "message": str(e),
            "error_code": "EXECUTION_ERROR"
        }
```

### 4. TDD 實施流程 (Red-Green-Refactor)

#### 第一步：Red (寫失敗測試)
```bash
# 首先創建失敗測試
python -m pytest tests/unit/test_[功能名稱].py -v  # 應該失敗
```

#### 單元測試結構 (先寫測試)
```python
import pytest
from unittest.mock import patch, Mock
from mcp_server_fastmcp import [功能名稱]

class Test[功能名稱]:
    def test_[功能名稱]_success(self):
        """測試成功情況 - 先寫這個失敗測試"""
        # Arrange
        expected_result = {"status": "success", "data": {...}}
        
        # Act
        result = [功能名稱](valid_params)
        
        # Assert
        assert result == expected_result
        
    def test_[功能名稱]_parameter_validation(self):
        """測試參數驗證 - 先寫這個失敗測試"""
        # Test invalid parameters
        result = [功能名稱](invalid_params)
        assert result["status"] == "error"
        assert "PARAMETER_ERROR" in result["error_code"]
        
    def test_[功能名稱]_error_handling(self):
        """測試錯誤處理 - 先寫這個失敗測試"""
        with patch('utility.util_autocad.UtilAutoCAD') as mock_autocad:
            mock_autocad.side_effect = Exception("AutoCAD error")
            result = [功能名稱](valid_params)
            assert result["status"] == "error"
```

#### 第二步：Green (最小實現)
```python
# 在 mcp_server_fastmcp.py 中實現最小代碼讓測試通過
@mcp.tool()
def [功能名稱]([參數]) -> Dict[str, Any]:
    """最小實現，僅讓測試通過"""
    return {"status": "success", "data": {}}  # 最簡單的通過實現
```

#### 第三步：Refactor (重構改善)
```python
# 重構為完整實現，保持測試通過
@mcp.tool()
def [功能名稱]([參數]) -> Dict[str, Any]:
    """完整實現，包含所有邏輯"""
    # 完整的實現邏輯
```

#### 整合測試 (TDD 方式)
- 先寫整合測試，確保失敗
- 實現功能讓整合測試通過
- 在 `tests/integration/` 目錄下創建整合測試

### 5. 驗證實施結果

#### 功能驗證
- [ ] 基本功能按 PRP 規格運作
- [ ] 所有參數正確處理
- [ ] 錯誤情況適當處理
- [ ] 回傳值格式正確

#### 品質驗證
- [ ] 所有測試通過
- [ ] 代碼覆蓋率達標
- [ ] 符合代碼風格
- [ ] 日誌記錄完整

#### 整合驗證
- [ ] 與現有 MCP 系統整合
- [ ] AutoCAD 連接正常
- [ ] 不影響現有功能
- [ ] 符合 CLAUDE.md 規範

### 6. 文檔更新
- 更新相關的技術文檔
- 添加使用範例
- 更新 README.md（如需要）

## 範例使用

### 執行 create_new_drawing 功能
```
/execute-prp create_new_drawing
```

這將會：
1. 讀取 `context_engineering/PRPs/create_new_drawing.md`
2. 在 `mcp_server_fastmcp.py` 中實施功能
3. 創建完整的測試套件
4. 運行所有測試確保品質
5. 驗證與現有系統的整合

### 執行 draw_line 功能
```
/execute-prp draw_line
```

這將為繪製直線功能進行完整的實施。

## 實施檢查清單

### 開發階段
- [ ] 讀取並理解完整的 PRP
- [ ] 實施核心功能代碼
- [ ] 添加完整的錯誤處理
- [ ] 實現參數驗證邏輯
- [ ] 添加詳細的日誌記錄

### 測試階段
- [ ] 創建單元測試
- [ ] 創建整合測試
- [ ] 執行所有測試
- [ ] 確保代碼覆蓋率達標
- [ ] 測試錯誤情況

### 整合階段
- [ ] 整合到現有 MCP 系統
- [ ] 測試與 AutoCAD 的連接
- [ ] 驗證與 Odoo 的相容性
- [ ] 確保不影響現有功能

### 驗收階段
- [ ] 功能按規格運作
- [ ] 所有測試通過
- [ ] 代碼品質達標
- [ ] 文檔完整準確

## 品質標準

### 代碼品質
- 遵循 PEP 8 代碼風格
- 使用類型提示
- 添加適當的註釋
- 遵循 CLAUDE.md 規範

### 測試品質
- 代碼覆蓋率 ≥ 95%
- 測試所有邊界情況
- 包含整合測試
- 測試執行時間合理

### 功能品質
- 完整的錯誤處理
- 清晰的錯誤訊息
- 適當的日誌記錄
- 良好的用戶體驗

## 常見問題處理

### AutoCAD 連接問題
- 確保 AutoCAD 正在運行
- 檢查 COM 註冊狀態
- 處理權限問題

### 測試問題
- 模擬 AutoCAD 環境
- 處理異步操作
- 確保測試隔離

### 整合問題
- 檢查現有功能影響
- 驗證 MCP 工具註冊
- 確保日誌系統一致

## 後續維護
- 定期更新文檔
- 監控功能性能
- 根據用戶反饋改進
- 保持代碼品質

---

**注意**: 執行此命令時，請確保完全理解 PRP 的所有要求，並嚴格按照規格實施。所有實施都必須通過完整的測試和驗證流程。