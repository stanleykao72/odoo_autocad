# AutoCAD-Odoo 整合系統 v5.0

> **AI助手整合版本** - 完整MCP支援與自然語言操作

Windows桌面應用程式，連接Odoo ERP與AutoCAD，提供AI驅動的工程/建築工作流程整合。

## 🚀 快速開始

```bash
# 啟動應用程式
python odoo.py

# 啟動AI助手伺服器模式
python odoo.py --mcp-server

# 測試MCP連接
python test_mcp_connection.py
```

## ✨ 主要功能

- **AutoCAD整合**: 自動提取圖面參數和BOQ資料
- **Odoo ERP整合**: 採購單處理和產品驗證
- **AI助手**: 7個MCP工具，支援自然語言操作
- **現代化UI**: CustomTkinter介面，圖示化控制

## 🤖 AI工具

**AutoCAD**: `scan_all_entities`, `get_table_data`, `extract_layout_info`, `query_entities_by_type`  
**Odoo**: `get_products_by_query`, `push_boq_to_project`, `validate_product_mapping`

## 📚 文檔

詳細文檔請參考 [`doc/`](./doc/) 目錄：

### 核心文檔
- **[MCP整合計劃](./doc/MCP_INTEGRATION_PLAN.md)** - 完整MCP實作文檔
- **[部署指南](./doc/DEPLOYMENT_GUIDE.md)** - 建置和部署說明
- **[發布說明](./doc/RELEASE_NOTES_v5.0.md)** - v5.0版本特色

### 設定與配置
- **[Gemini CLI設定](./doc/GEMINI_CLI_SETUP.md)** - AI助手配置指南
- **[開發指南](./doc/README-DEVELOPMENT.md)** - 開發環境設定

### 安全與部署
- **[防毒軟體解決方案](./doc/ANTIVIRUS_SOLUTION.md)** - 防毒軟體相容性指南
- **[程式碼簽章指南](./doc/CODE_SIGNING_GUIDE.md)** - 數位簽章設定
- **[內部CA指南](./doc/INTERNAL_CA_GUIDE.md)** - 企業憑證管理

### UI與介面
- **[UI改善計劃](./doc/UI_IMPROVEMENT_PLAN.md)** - 介面現代化計劃

### 進階整合
- **[AutoCAD MCP整合方案](./doc/AutoCAD_MCP_Integration_Plan.md)** - 第三方AutoCAD工具整合計劃
- **[Context Engineering分析](./doc/Context_Engineering_Analysis.md)** - AI協作開發方法論應用

### Context Engineering 工作流程
- **[初始需求](./INITIAL.md)** - 功能需求定義
- **[Context Engineering框架](./context_engineering/README.md)** - 開發方法論實施
- **自定義命令**: `/generate-prp` 和 `/execute-prp` 用於標準化開發流程

## 🛠️ 開發

此專案遵循TDD方法論，詳細開發指引請參考 [`CLAUDE.md`](./CLAUDE.md)。

```bash
# 安裝依賴
pip install -r requirements.txt

# 執行測試
python -m pytest tests/

# 建置EXE
python build_exe.py
```

## 📞 支援

- **GitHub**: [stanleykao72/odoo_autocad](https://github.com/stanleykao72/odoo_autocad)
- **文檔**: [`doc/`](./doc/) 目錄包含完整說明
- **版本**: 5.0 (MCP整合完整版)

---

*整合AutoCAD與Odoo的智能工程解決方案*