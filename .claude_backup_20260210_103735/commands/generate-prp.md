# /generate-prp Command

## 目的
根據 INITIAL.md 中的功能需求，為指定的功能生成詳細的 PRP (Product Requirements Prompt)。

## 使用方法
```
/generate-prp [功能名稱]
```

## 參數
- `[功能名稱]`: 要生成 PRP 的功能名稱（例如：create_new_drawing, draw_line, set_layer）

## 執行步驟

### 1. 讀取基礎資訊
- 閱讀 `INITIAL.md` 了解整體需求
- 查看 `CLAUDE.md` 了解開發規範
- 參考 `context_engineering/templates/prp_template.md` 模板

### 2. 分析功能需求
- 從 INITIAL.md 中提取指定功能的具體需求
- 分析功能的輸入參數和輸出格式
- 確定與現有系統的整合點

### 3. 生成 PRP 內容
基於 PRP 模板，生成包含以下內容的詳細 PRP：

#### 功能概述
- 目標和用途
- 使用場景
- 成功標準

#### 技術規格
- 函數簽名
- 參數規格
- 回傳值規格

#### 實施要求
- 依賴項目
- 核心邏輯
- 錯誤處理

#### 範例代碼
- 基本實現
- 使用範例
- 測試代碼

#### 驗證標準
- 功能驗證
- 品質標準
- 整合標準

### 4. 整合現有系統
- 確保與現有 MCP 工具的一致性
- 使用現有的 `UtilAutoCAD` 和 `UtilOdoo` 類別
- 遵循現有的錯誤處理模式

#### 系統修改特殊考量
對於修改既有系統（如 `mcp_server_fastmcp.py`），額外包含：
- **向後相容性分析**: 確保新功能不影響既有功能
- **既有模式遵循**: 使用相同的裝飾器、錯誤處理、日誌格式
- **整合點識別**: 明確指出修改位置和整合方式
- **回歸測試**: 確保既有功能仍正常運作

### 5. 輸出 PRP 檔案
- 將生成的 PRP 儲存到 `context_engineering/PRPs/[功能名稱].md`
- 確保格式符合模板要求
- 包含所有必要的技術細節

## 範例使用

### 生成 create_new_drawing 功能的 PRP
```
/generate-prp create_new_drawing
```

這將會：
1. 分析 INITIAL.md 中關於創建新圖面的需求
2. 生成完整的 PRP 文檔
3. 包含詳細的實施代碼範例
4. 定義測試要求和驗證標準
5. 儲存到 `context_engineering/PRPs/create_new_drawing.md`

### 生成 draw_line 功能的 PRP
```
/generate-prp draw_line
```

這將為繪製直線功能生成詳細的 PRP。

## 輸出格式
生成的 PRP 將包含：
- 完整的功能描述
- 詳細的技術規格
- 實際可執行的代碼範例
- 完整的測試套件
- 清晰的驗證標準

## 注意事項
- 確保生成的 PRP 符合 Context Engineering 標準
- 所有代碼範例都應該是實際可執行的
- 必須包含完整的錯誤處理
- 遵循現有的代碼風格和架構模式

## 後續步驟
生成 PRP 後，可以使用 `/execute-prp [功能名稱]` 命令來實際實施功能。