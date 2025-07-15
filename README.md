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

- **[MCP整合計劃](./doc/MCP_INTEGRATION_PLAN.md)** - 完整MCP實作文檔
- **[部署指南](./doc/DEPLOYMENT_GUIDE.md)** - 建置和部署說明
- **[Gemini CLI設定](./doc/GEMINI_CLI_SETUP.md)** - AI助手配置指南
- **[發布說明](./doc/RELEASE_NOTES_v5.0.md)** - v5.0版本特色
- **[開發指南](./doc/README-DEVELOPMENT.md)** - 開發環境設定

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