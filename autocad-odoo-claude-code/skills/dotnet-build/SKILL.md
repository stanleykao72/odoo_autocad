# .NET 建置 Skill

---
name: dotnet-build
description: dotnet CLI 建置、測試、發佈指令與 NuGet 套件管理
trigger-keywords:
  - dotnet build
  - 建置
  - 編譯
  - NuGet
  - csproj
  - 發佈
  - publish
  - restore
allowed-tools:
  - Bash
  - Read
  - Edit
  - Glob
---

## 概述

本專案的 `dotnet` CLI 不在 PATH 中，必須使用完整路徑。

## 核心指令

### 建置

```bash
# 完整解決方案建置
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln

# 單一專案建置
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/src/OdooAutoCAD.App/OdooAutoCAD.App.csproj

# 清理重建
"C:\Program Files\dotnet\dotnet.exe" clean csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln && \
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln
```

### 測試

```bash
# 執行所有測試
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --verbosity normal

# 特定測試
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --filter "FullyQualifiedName~ShutdownCleanup"
```

### NuGet 套件管理

```bash
# 新增套件
"C:\Program Files\dotnet\dotnet.exe" add csharp/OdooAutoCADIntegration/src/OdooAutoCAD.App/OdooAutoCAD.App.csproj package PackageName

# 還原套件
"C:\Program Files\dotnet\dotnet.exe" restore csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln
```

## 已知警告（可忽略）

| 專案 | 警告 | 說明 |
|------|------|------|
| Core | CS8602 | null dereference（既有） |
| MCP | CS1998 | async method without await（既有） |

**規則**: 不引入新的建置警告。如果建置產生新警告，必須修復。

## .csproj 修改規則

- 不擅自更新 `<Version>` 標籤
- 不移除現有的 `<PackageReference>`
- 新增套件前確認必要性
- Integration.Tests 專案需要 `<UseWPF>true</UseWPF>`

## 建置驗證流程

每次修改程式碼後：
1. 執行 `dotnet build` 確認編譯通過
2. 檢查無新增警告
3. 執行相關測試確認無回歸
