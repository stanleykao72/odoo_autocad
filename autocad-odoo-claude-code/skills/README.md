# Skills 系統

## 概述

Skills 是自動觸發的專業功能模組。當對話中出現特定關鍵字時，對應的 Skill 會被自動載入，提供專業的開發指引。

## 可用 Skills

| Skill | 觸發關鍵字 | 說明 |
|-------|-----------|------|
| `csharp-mvvm` | ObservableProperty, RelayCommand, ViewModel | CommunityToolkit.Mvvm 模式 |
| `xaml-patterns` | XAML, DataTemplate, Style, Converter | WPF XAML 開發模式 |
| `dotnet-build` | dotnet build, 建置, NuGet, csproj | .NET 建置與套件管理 |
| `autocad-com` | COM, AutoCAD, Interop, Marshal | AutoCAD COM Interop |
| `odoo-rest-api` | Odoo API, REST, 同步, 認證 | Odoo REST API 整合 |
| `git-workflow` | git commit, 分支, merge, PR | Git 操作規範 |
| `wpf-navigation` | NavigationService, 導航, 頁面切換 | 導航系統機制 |
| `di-patterns` | DI, 依賴注入, Hosting | 依賴注入配置 |
| `i18n-cjk` | CJK, 字型, 中文, 本地化 | CJK 國際化支援 |
| `debug-patterns` | 除錯, debug, Exception, crash | 除錯診斷策略 |
| `python-legacy` | Python, 舊程式, 遷移, legacy | Python 舊系統分析 |

## Skill 檔案格式

每個 Skill 目錄包含一個 `SKILL.md` 檔案，使用 YAML frontmatter 定義元資料：

```yaml
---
name: skill-name
description: 簡要說明
trigger-keywords:
  - keyword1
  - keyword2
allowed-tools:
  - Read
  - Write
  - Edit
---
```

## 新增 Skill 步驟

1. 在 `skills/` 下建立新目錄
2. 建立 `SKILL.md` 檔案
3. 定義 YAML frontmatter（name, description, trigger-keywords, allowed-tools）
4. 撰寫 Skill 內容（模式、範例、檢查清單）
5. 更新本 README 的 Skills 表格
6. 在根 CLAUDE.md 中加入觸發關鍵字映射
