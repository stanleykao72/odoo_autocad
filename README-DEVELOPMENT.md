# 開發環境設置說明

## 環境需求

### Windows 生產環境
- Python 3.10+
- AutoCAD (支援COM API)
- Odoo服務器連接

### macOS/Linux 開發環境 
- Python 3.10+
- 僅支援UI開發和API測試

## 安裝步驟

### 1. 創建conda環境

```bash
# 創建環境
conda create -n odoo_autocad python=3.12

# 激活環境
conda activate odoo_autocad

# 安裝tkinter (macOS/Linux)
conda install tk
```

### 2. 安裝依賴

**Windows環境:**
```bash
pip install -r requirements-windows.txt
```

**macOS/Linux開發環境:**
```bash
pip install -r requirements.txt
```

## 運行應用

### Windows完整功能
```bash
python odoo.py
```

### macOS/Linux UI測試
```bash
# 基礎跨平台UI測試
python test_ui_cross_platform.py

# 完整UI功能演示
python test_enhanced_ui.py
```

## 文件結構

```
odoo_autocad/
├── config/                    # 配置文件
├── db/                       # SQLite數據庫
├── forms/                    # UI表單
│   ├── form_main.py         # 原始主表單 (Windows)
│   └── form_main_modern.py  # 現代化主表單
├── models/                   # 數據模型
├── ui/                       # UI組件和主題
│   ├── ui_theme.py          # 主題配置
│   ├── ui_fonts.py          # 字體管理
│   └── enhanced_widgets.py  # 增強UI組件
├── utility/                  # 工具模組
├── test_*.py                # 測試文件
├── requirements.txt         # 基礎依賴
├── requirements-windows.txt # Windows專用依賴
└── odoo.py                  # 主程式入口
```

## 開發注意事項

1. **跨平台相容性**: Windows專用功能（AutoCAD COM）在其他平台無法使用
2. **UI測試**: 使用test_*.py文件進行UI開發和測試
3. **字體支援**: 已實現跨平台字體管理系統
4. **主題系統**: 支援亮色/暗色主題切換

## 已知問題

- macOS/Linux環境無法使用AutoCAD COM功能
- 部分Windows特定依賴在其他平台不可用
- 完整功能需要Windows + AutoCAD環境

## UI現代化功能

- ✅ CustomTkinter現代化UI框架
- ✅ 響應式佈局設計
- ✅ 跨平台字體管理
- ✅ 動態狀態指示器
- ✅ 增強的搜尋和過濾功能
- ✅ 實時日誌查看器
- ✅ 進度對話框和載入狀態
- ✅ 主題和外觀自定義