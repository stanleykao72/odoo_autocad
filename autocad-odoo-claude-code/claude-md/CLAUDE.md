# CLAUDE.md — AutoCAD-Odoo 整合桌面應用

> **版本**: 1.0
> **適用**: C# WPF 桌面應用 (Odoo ERP ↔ AutoCAD 整合)
> **語言**: 正體中文為主，技術術語保留英文

---

## 溝通協議

- 直接、精確，不做推測
- 使用正體中文溝通，技術術語保留英文
- 遵循 **DO NOT OVERDESIGN** 原則：只做被要求的事
- 修改程式碼前必須先讀取檔案
- 不主動建立不必要的文件

---

## 專案概述

Windows 桌面應用程式，橋接 Odoo ERP 與 AutoCAD：
- **前端**: WPF + CommunityToolkit.Mvvm (MVVM 架構)
- **後端**: .NET 8, Microsoft.Extensions.Hosting (DI)
- **AutoCAD 整合**: COM Interop (未來)
- **Odoo 整合**: REST API Client (未來)
- **資料層**: SQLite + Entity Framework Core (未來)

---

## 核心規則

### 絕對禁止
- **禁止修改** NuGet 套件原始碼或 .NET 框架程式碼
- **禁止** `git add .` 或 `git add -A`（使用具體檔案名稱）
- **禁止** `git push --force`、`git reset --hard`、`git checkout .`
- **禁止** 在未讀取檔案前修改它
- **禁止** 擅自更新 `.csproj` 版本號（除非明確要求）
- **禁止** 引入不必要的 NuGet 套件

### 必須遵守
- 使用 `"C:\Program Files\dotnet\dotnet.exe"` 執行 dotnet 命令（不在 PATH 中）
- 每次修改後執行建置驗證
- 遵循 CommunityToolkit.Mvvm 模式（`[ObservableProperty]`、`[RelayCommand]`）
- 所有頁面放在 `Views/Pages/{Name}Page.xaml`
- ViewModel 放在 `ViewModels/{Name}ViewModel.cs`
- 使用 DI 注入服務，不使用 `new` 直接建立

---

## 多代理協作系統

本專案使用 6 個專業化代理（Agents），遵循 **SDD (Specification-Driven Development)** 方法論。

### 可用代理

| 代理 | 職責 | 觸發條件 |
|------|------|----------|
| WPF 開發者 | XAML/ViewModel/UI 開發 | UI 元件、頁面、樣式相關任務 |
| AutoCAD 整合專家 | COM Interop、圖檔操作 | AutoCAD 連線、參數提取相關 |
| Odoo API 開發者 | REST API、資料同步 | Odoo 連線、資料同步相關 |
| QA 測試工程師 | 單元/整合測試 | 測試、品質保證相關 |
| 專案協調者 | 任務分解、排程管理 | 跨代理任務、整體規劃 |
| 安全審查員 | 安全漏洞、權限檢查 | 安全審查、程式碼審計 |

### 代理調度規則
- 依據 `rules/agents.md` 自動選擇適當代理
- 使用 `templates/shared_handoffs/` 進行交接
- 任何跨代理任務需由專案協調者統籌

---

## 階層式 CLAUDE.md 結構

```
CLAUDE.md (根層級 — 全域規則)
└── csharp/CLAUDE.md (C# WPF 專案 — 開發規範)
```

- **根層級**: 專案架構、核心規則、代理系統、工具使用
- **csharp 層級**: C# 編碼標準、WPF 模式、建置指令

---

## Skill 觸發系統

當對話中出現特定關鍵字時，自動載入對應 Skill：

| Skill | 觸發關鍵字 |
|-------|-----------|
| `csharp-mvvm` | ObservableProperty, RelayCommand, ViewModel, MVVM |
| `xaml-patterns` | XAML, DataTemplate, Style, Converter, Binding |
| `dotnet-build` | dotnet build, 建置, 編譯, NuGet, csproj |
| `autocad-com` | COM, AutoCAD, Interop, Drawing |
| `odoo-rest-api` | Odoo API, REST, 同步, 認證 |
| `git-workflow` | git commit, 分支, merge, PR |
| `wpf-navigation` | NavigationService, 導航, 頁面切換 |
| `di-patterns` | DI, 依賴注入, IServiceProvider, Hosting |
| `i18n-cjk` | CJK, 字型, 中文, 本地化, 翻譯 |
| `debug-patterns` | 除錯, debug, 例外, crash, 堆疊追蹤 |

---

## 文件規範

### 文件模板
所有開發文件使用 `templates/` 中的標準模板：
- **US-xxx**: 用戶故事（User Story）
- **TASK-BE-xxx**: 後端任務（Backend Task）
- **TASK-FE-xxx**: 前端任務（Frontend Task）
- **TS-xxx**: 測試規格（Test Specification）
- **BUG-xxx**: 缺陷報告（Bug Report）
- **FR-xxx**: 功能需求（Feature Request）

### 文件撰寫標準
- 使用正體中文撰寫
- 使用 Mermaid 圖表描述流程
- 使用表格整理結構化資訊
- 程式碼範例使用對應語言的語法高亮

---

## 建置與測試指令

```bash
# 建置
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln

# 測試
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --verbosity normal

# 在 bash 中複製檔案使用 cp（非 copy）
cp source dest
```

---

## 工具使用要求

- 使用 `Read` 讀取檔案（非 `cat`/`head`/`tail`）
- 使用 `Edit` 編輯檔案（非 `sed`/`awk`）
- 使用 `Write` 建立檔案（非 `echo`/`cat <<EOF`）
- 使用 `Glob` 搜尋檔案（非 `find`/`ls`）
- 使用 `Grep` 搜尋內容（非 `grep`/`rg`）
- 使用 `Task` 工具委派複雜任務給子代理
