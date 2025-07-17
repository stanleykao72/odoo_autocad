# Context Engineering 框架

> **目標**: 使用 Context Engineering 方法論開發 AutoCAD MCP 功能  
> **版本**: v1.0  
> **建立日期**: 2025年7月16日  

## 框架結構

```
context_engineering/
├── PRPs/                  # Product Requirements Prompts
│   ├── create_drawing.md  # 新建圖面功能 PRP
│   ├── draw_line.md       # 繪製直線功能 PRP
│   └── ...
├── examples/              # 代碼範例庫
│   ├── mcp_tool_template.py
│   ├── autocad_com_examples.py
│   └── error_handling_patterns.py
├── validation/            # 驗證腳本
│   ├── test_templates.py
│   └── quality_checks.py
├── templates/             # 模板文件
│   ├── mcp_tool_template.py
│   ├── prp_template.md
│   └── test_template.py
└── README.md             # 本文件
```

## 使用流程

### 標準 Context Engineering 流程
1. **準備階段**: 閱讀 `INITIAL.md` 了解整體需求
2. **生成 PRP**: 使用 `/generate-prp [功能名稱]` 生成詳細規格
3. **執行 PRP**: 使用 `/execute-prp [功能名稱]` 實施功能
4. **驗證結果**: 確保功能符合 PRP 要求

### TDD + Context Engineering 流程 (推薦)
1. **準備階段**: 閱讀 `INITIAL.md` 和 `TDD_CONTEXT_ENGINEERING_GUIDE.md`
2. **生成 PRP**: 使用 `/generate-prp [功能名稱]` 生成包含 TDD 要求的 PRP
3. **TDD 執行**: 使用 `/execute-prp [功能名稱]` 執行 Red-Green-Refactor 循環
4. **測試驗證**: 確保 95%+ 測試覆蓋率和所有測試通過

### 手動流程（傳統方式）
1. **功能需求定義**
   - 複製 `templates/prp_template.md` 到 `PRPs/` 目錄
   - 根據具體功能需求填寫 PRP 模板
   - 添加相關的代碼範例到 `examples/` 目錄

2. **實施階段**
   - 使用完整的 PRP 作為 AI 助手的上下文
   - 基於範例代碼進行實施
   - 遵循 CLAUDE.md 中的全局規則

3. **驗證階段**
   - 執行驗證腳本檢查代碼品質
   - 運行單元測試和整合測試
   - 確保符合功能需求

## 快速開始

### 創建新功能
```bash
# 1. 複製 PRP 模板
cp templates/prp_template.md PRPs/my_new_feature.md

# 2. 編輯 PRP 內容
# 在 PRPs/my_new_feature.md 中填寫詳細需求

# 3. 使用 PRP 作為 AI 助手的上下文進行開發
```

### 範例使用
參考 `PRPs/create_drawing.md` 了解如何撰寫完整的 PRP。

## 最佳實踐

### Context Engineering 最佳實踐
1. **詳細描述**: PRP 應該包含完整的功能描述和技術要求
2. **提供範例**: 每個功能都應該有相應的代碼範例
3. **驗證標準**: 明確定義成功標準和驗證方法
4. **持續更新**: 根據實施經驗更新範例和模板

### TDD 最佳實踐
1. **先寫測試**: 總是先寫失敗測試，然後實現功能
2. **最小實現**: 實現最小的代碼讓測試通過
3. **持續重構**: 保持測試通過的同時改善代碼品質
4. **高覆蓋率**: 確保代碼覆蓋率達到 95% 以上

### 整合最佳實踐
1. **PRP 包含 TDD**: 每個 PRP 都應包含詳細的 TDD 測試要求
2. **測試驅動設計**: 使用測試來驅動 API 設計
3. **持續驗證**: 在每個階段都要運行測試
4. **文檔即代碼**: 測試作為功能的活文檔

## 注意事項

- 所有代碼都應該遵循 `CLAUDE.md` 中的全局規則
- 使用現有的 `UtilAutoCAD` 和 `UtilOdoo` 類別
- 保持與現有 MCP 工具的一致性
- 確保完整的錯誤處理和日誌記錄
- **TDD 強制要求**: 所有新功能都必須先寫測試
- **測試覆蓋率**: 核心邏輯必須達到 95% 以上覆蓋率

## 相關文檔

- [TDD + Context Engineering 整合指南](./TDD_CONTEXT_ENGINEERING_GUIDE.md)
- [CLAUDE.md 開發規範](../CLAUDE.md)
- [INITIAL.md 功能需求](../INITIAL.md)