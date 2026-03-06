# C# v6.0 專案目錄說明

> **版本**: 6.0.0 (Sprint 2)  
> **最後更新**: 2026-02-10

## 目錄結構

```
csharp/
├── OdooAutoCADIntegration/     # 原始碼專案
│   ├── src/
│   │   ├── OdooAutoCAD.App/           # WPF 應用程式
│   │   ├── OdooAutoCAD.Core/          # 核心業務邏輯
│   │   ├── OdooAutoCAD.Threading/     # GUI Proxy 系統
│   │   ├── OdooAutoCAD.MCP/           # MCP Server
│   │   ├── OdooAutoCAD.Data/          # 資料層
│   │   └── OdooAutoCAD.Configuration/ # 配置管理
│   └── tests/                         # 測試專案 (待建立)
│
├── publish/                    # 編譯輸出目錄 (~200 MB)
│   ├── OdooAutoCAD.exe
│   ├── appsettings.json
│   └── [所有相依套件]
│
└── installer/                  # 📦 建置與安裝檔案
    ├── build_and_package.bat           # 完整建置腳本 (批次檔)
    ├── build_and_package.ps1           # 完整建置腳本 (PowerShell)
    ├── build_installer_only.bat        # 快速建立安裝檔 ⭐
    ├── odoo-autocad-csharp-setup.iss   # Inno Setup 腳本
    ├── odoo-autocad-integration-6.0.0-setup.exe  # 安裝程式
    ├── BUILD_GUIDE.md                  # 詳細建置指南
    ├── QUICK_BUILD_SPRINT2.md          # 快速建置說明
    └── SPRINT2_BUILD_SUMMARY.md        # 建置總結
```

---

## 快速開始

### 建立安裝檔

```cmd
cd C:\odoo\autocad_source\csharp\installer
build_installer_only.bat
```

### 完整建置

```cmd
cd C:\odoo\autocad_source\csharp\installer
build_and_package.bat
```

---

## 文件說明

### 📖 必讀文件

1. **BUILD_GUIDE.md** - 完整的建置指南
   - 系統需求
   - 建置步驟
   - 常見問題
   - 測試建議

2. **QUICK_BUILD_SPRINT2.md** - 快速建置說明
   - 快速上手
   - 常見錯誤解決

3. **SPRINT2_BUILD_SUMMARY.md** - Sprint 2 總結
   - 已完成功能
   - 建置選項
   - 測試檢查清單

### 📜 原始碼文件

- `OdooAutoCADIntegration/README.md` - 專案技術說明

### 📋 需求文件

- `../docs/requirements/index.md` - 需求規格總索引
- `../docs/requirements/master-task-sequence.md` - 500個任務排程

---

## Sprint 2 成果

### ✅ 已實作 (64/500 任務)
- 主視窗 + 側邊欄導航
- 設定頁面 (Odoo/AutoCAD/MCP)
- GUI Proxy 系統
- MCP Server 框架
- 應用程式資訊顯示

### 📦 可交付成果
- 安裝程式: `installer/odoo-autocad-integration-6.0.0-setup.exe`
- 大小: ~220 MB (包含 .NET 8 Runtime)
- 支援: Windows 10/11 (64-bit)

---

## 下一步

**Sprint 3** - AutoCAD & Odoo Core Connections (40 tasks)
- AutoCAD COM 整合
- Odoo REST API 整合
- Dashboard 功能頁面

---

## 支援

如有問題，請參考:
- 建置問題: `installer/BUILD_GUIDE.md`
- 專案架構: `OdooAutoCADIntegration/README.md`
- 需求規格: `../docs/requirements/`
