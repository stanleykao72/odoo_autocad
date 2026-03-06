
# Sprint 2 快速建置指南

## ✅ 系統環境說明
**PowerShell Core (pwsh) 已安裝完成！** 版本：7.5.4

現在可以使用自動化建置腳本了。

## 🚀 建置方式

### 選項 A: 使用 PowerShell Core 自動建置 (推薦) ⭐

1. **開啟 PowerShell Core**:
```powershell
pwsh
```

2. **執行自動建置腳本**:
```powershell
cd C:\odoo\autocad_source\csharp
.\build_and_package.ps1
```

3. **等待建置完成** (約 3-5 分鐘)

---

### 選項 B: 使用命令提示字元 (CMD) 手動執行

1. **開啟命令提示字元** (以系統管理員身分執行)

2. **執行建置腳本**:
```cmd
cd C:\odoo\autocad_source\csharp
build_and_package.bat
```

3. **等待建置完成** (約 3-5 分鐘)

### 選項 C: 手動步驟建置

如果批次檔無法執行，請按照以下步驟手動建置：

#### Step 1: 清理舊檔案
```cmd
cd C:\odoo\autocad_source\csharp
if exist publish rmdir /s /q publish
```

#### Step 2: 進入專案目錄並還原套件
```cmd
cd OdooAutoCADIntegration
dotnet restore
```

#### Step 3: 建置發行版本
```cmd
dotnet publish src\OdooAutoCAD.App\OdooAutoCAD.App.csproj ^
  --configuration Release ^
  --runtime win-x64 ^
  --self-contained true ^
  --output ..\publish ^
  /p:PublishSingleFile=false ^
  /p:PublishReadyToRun=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  /p:DebugType=none ^
  /p:DebugSymbols=false
```

#### Step 4: 建立安裝檔
```cmd
cd ..
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\odoo-autocad-csharp-setup.iss
```

---

### 選項 D: 使用 Windows PowerShell (非 Core)

1. **開啟 Windows PowerShell** (不是 PowerShell Core)

2. **執行以下指令**:
```powershell
# 允許執行腳本 (僅需執行一次)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 切換目錄
Set-Location C:\odoo\autocad_source\csharp

# 執行建置
.\build_and_package.ps1
```

---

## 預期輸出

### 建置成功後，您會看到:

```
========================================
BUILD COMPLETED SUCCESSFULLY
========================================

Output Files:
   Published Files: C:\odoo\autocad_source\csharp\publish
   Installer:       C:\odoo\autocad_source\csharp\installer\odoo-autocad-integration-6.0.0-setup.exe

Installer Size: ~220,000,000 bytes (約 220 MB)
```

### 檔案位置:
- **執行檔**: `C:\odoo\autocad_source\csharp\publish\OdooAutoCAD.exe`
- **安裝檔**: `C:\odoo\autocad_source\csharp\installer\odoo-autocad-integration-6.0.0-setup.exe`

---

## 目前已存在的檔案

根據檢查，`publish` 目錄**已經有編譯好的檔案**，包括:
- ✅ OdooAutoCAD.exe (主程式)
- ✅ 所有相依 DLL
- ✅ .NET 8 Runtime

您可以:
1. **直接執行測試**: `publish\OdooAutoCAD.exe`
2. **直接建立安裝檔** (如果 publish 檔案是最新的):
   ```cmd
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\odoo-autocad-csharp-setup.iss
   ```

---

## 驗證建置

### 測試執行檔:
```cmd
cd C:\odoo\autocad_source\csharp\publish
OdooAutoCAD.exe
```

### 檢查版本資訊:
```cmd
cd C:\odoo\autocad_source\csharp\publish
OdooAutoCAD.exe --version
```

---

## 如果遇到問題

### 問題 1: dotnet 命令找不到
**解決方式**: 安裝 .NET 8 SDK
- 下載: https://dotnet.microsoft.com/download/dotnet/8.0
- 安裝後重新開啟命令提示字元

### 問題 2: Inno Setup 找不到
**解決方式**: 確認安裝路徑
```cmd
dir "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
```
如果不存在，請從 https://jrsoftware.org/isdl.php 下載安裝

### 問題 3: 編譯錯誤
**解決方式**: 檢查錯誤訊息
- 如果是缺少套件，執行 `dotnet restore`
- 如果是語法錯誤，查看錯誤訊息中的檔案和行號

---

## 快速檢查清單

建置前請確認:
- [ ] 已安裝 .NET 8 SDK (`dotnet --version`)
- [ ] 已安裝 Inno Setup 6
- [ ] 有足夠磁碟空間 (至少 2 GB)
- [ ] 以系統管理員身分執行命令提示字元

建置後請檢查:
- [ ] `publish\OdooAutoCAD.exe` 存在
- [ ] `installer\odoo-autocad-integration-6.0.0-setup.exe` 存在
- [ ] 安裝檔大小約 200-250 MB
- [ ] 可以正常執行 `publish\OdooAutoCAD.exe`

---

## 建議的執行順序 (今天)

由於無法自動執行腳本，建議:

1. **先測試現有的執行檔**:
   ```cmd
   C:\odoo\autocad_source\csharp\publish\OdooAutoCAD.exe
   ```

2. **如果運作正常，直接建立安裝檔**:
   ```cmd
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" C:\odoo\autocad_source\csharp\installer\odoo-autocad-csharp-setup.iss
   ```

3. **測試安裝檔**:
   - 執行 `installer\odoo-autocad-integration-6.0.0-setup.exe`
   - 安裝到測試路徑
   - 驗證功能

---

**完成時間預估**: 5-10 分鐘 (如果 publish 檔案已是最新)
