# 完整建置和封裝流程腳本 v5.0 (PowerShell版本)
# 從原始碼到安裝包的一鍵建置

Write-Host "🚀 Odoo-AutoCAD 整合系統 v5.0 - 完整建置流程" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""

# 檢查必要工具
Write-Host "📋 檢查建置環境..." -ForegroundColor Cyan

# 檢查 Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python: $pythonVersion" -ForegroundColor Yellow
} catch {
    Write-Host "❌ Python 未安裝或不在 PATH 中" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出..."
    exit 1
}

# 檢查 PyInstaller
try {
    $pyinstallerVersion = pyinstaller --version 2>&1
    Write-Host "✅ PyInstaller: $pyinstallerVersion" -ForegroundColor Yellow
} catch {
    Write-Host "❌ PyInstaller 未安裝，請執行: pip install pyinstaller" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出..."
    exit 1
}

# 檢查 Inno Setup
$innoPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $innoPath)) {
    Write-Host "❌ Inno Setup 6 未安裝，請先安裝 Inno Setup" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出..."
    exit 1
} else {
    Write-Host "✅ Inno Setup 6 已安裝" -ForegroundColor Yellow
}

Write-Host "✅ 建置環境檢查完成" -ForegroundColor Green
Write-Host ""

# 步驟 1: 建置 EXE
Write-Host "🔨 步驟 1/3: 建置可執行檔..." -ForegroundColor Cyan
try {
    # 直接使用 PyInstaller 建置，不依賴 build_windows.ps1
    & python -m PyInstaller --onefile --windowed --name "odoo-autocad-integration" --icon "icon/odoo_autocad.ico" --add-data "config;config" --add-data "fonts;fonts" --add-data "icon;icon" --add-data "db;db" --hidden-import "customtkinter" --hidden-import "win32com.client" --hidden-import "win32com.gen_py" --hidden-import "pywintypes" --hidden-import "win32api" --hidden-import "tkinter" --hidden-import "tkinter.ttk" --hidden-import "sqlalchemy" --hidden-import "sqlalchemy.ext.declarative" --hidden-import "sqlalchemy.orm" --exclude-module "pytest" --exclude-module "unittest" --exclude-module "doctest" --exclude-module "pdb" --exclude-module "matplotlib" --exclude-module "numpy" --exclude-module "pandas" --clean --noconfirm --distpath "output" odoo.py
    if ($LASTEXITCODE -eq 0) {
        # 複製額外檔案到輸出目錄
        Copy-Item "doc\ANTIVIRUS_SOLUTION.md" "output\" -ErrorAction SilentlyContinue
        Copy-Item "README.md" "output\" -ErrorAction SilentlyContinue
        Copy-Item "doc\README-DEVELOPMENT.md" "output\" -ErrorAction SilentlyContinue
        Write-Host "✅ EXE 建置完成" -ForegroundColor Green
    } else {
        throw "PyInstaller 回傳錯誤代碼: $LASTEXITCODE"
    }
} catch {
    Write-Host "❌ EXE 建置失敗: $_" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出..."
    exit 1
}
Write-Host ""

# 步驟 2: 程式碼簽章 (如果有憑證)
Write-Host "🔐 步驟 2/3: 程式碼簽章..." -ForegroundColor Cyan
$certPath = "certs\codesign.pfx"
$signtoolPath = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"

if (Test-Path $certPath) {
    Write-Host "正在簽署 output\odoo-autocad-integration.exe..." -ForegroundColor Yellow
    try {
        & $signtoolPath sign /f $certPath /p "YourSecurePassword123!" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "output\odoo-autocad-integration.exe"
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ 程式碼簽章完成" -ForegroundColor Green
        } else {
            Write-Host "⚠️ 程式碼簽章失敗，但繼續建置" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "⚠️ 簽章過程發生錯誤: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "ℹ️ 未找到憑證檔案 $certPath，跳過簽章" -ForegroundColor Gray
}
Write-Host ""

# 步驟 3: 建立安裝包
Write-Host "📦 步驟 3/3: 建立 Inno Setup 安裝包..." -ForegroundColor Cyan
try {
    # 使用無簽章版本以避免憑證密碼問題
    & $innoPath "installer\odoo-autocad-setup.iss"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ 安裝包建置完成" -ForegroundColor Green
    } else {
        throw "Inno Setup 回傳錯誤代碼: $LASTEXITCODE"
    }
} catch {
    Write-Host "❌ 安裝包建置失敗: $_" -ForegroundColor Red
    Read-Host "按 Enter 鍵退出..."
    exit 1
}
Write-Host ""

# 顯示完成結果
Write-Host "🎉 建置完成！" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host "📁 輸出檔案位置:" -ForegroundColor Cyan

# EXE 檔案資訊
$exePath = "output\odoo-autocad-integration.exe"
if (Test-Path $exePath) {
    $exeInfo = Get-Item $exePath
    Write-Host "   EXE: $exePath" -ForegroundColor White
    Write-Host "   📊 EXE 大小: $($exeInfo.Length.ToString('N0')) bytes" -ForegroundColor White
    
    # 計算 SHA256
    $hash = Get-FileHash -Path $exePath -Algorithm SHA256
    Write-Host "   🔒 SHA256: $($hash.Hash.Substring(0,16))..." -ForegroundColor White
}

# 安裝包檔案資訊
$installerPath = "installer\odoo-autocad-integration-5.0-setup.exe"
if (Test-Path $installerPath) {
    $installerInfo = Get-Item $installerPath
    Write-Host "   安裝包: $installerPath" -ForegroundColor White
    Write-Host "   📊 安裝包大小: $($installerInfo.Length.ToString('N0')) bytes" -ForegroundColor White
}

Write-Host ""
Write-Host "💡 建議:" -ForegroundColor Yellow
Write-Host "   1. 測試安裝包在乾淨系統上的安裝" -ForegroundColor White
Write-Host "   2. 檢查防毒軟體相容性" -ForegroundColor White  
Write-Host "   3. 驗證數位簽章狀態" -ForegroundColor White
Write-Host ""

Read-Host "按 Enter 鍵關閉..."