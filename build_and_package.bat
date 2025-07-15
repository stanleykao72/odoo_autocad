@echo off
REM 完整建置和封裝流程腳本 v5.0
REM 從原始碼到安裝包的一鍵建置

echo 🚀 Odoo-AutoCAD 整合系統 v5.0 - 完整建置流程
echo ===============================================
echo.

REM 檢查必要工具
echo 📋 檢查建置環境...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Python 未安裝或不在 PATH 中
    pause
    exit /b 1
)

pyinstaller --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ PyInstaller 未安裝，請執行: pip install pyinstaller
    pause
    exit /b 1
)

if not exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo ❌ Inno Setup 6 未安裝，請先安裝 Inno Setup
    pause
    exit /b 1
)

echo ✅ 建置環境檢查完成
echo.

REM 步驟 1: 建置 EXE
echo 🔨 步驟 1/3: 建置可執行檔...
REM 直接使用 PyInstaller 建置，不依賴 build_windows.bat
python -m PyInstaller --onefile --windowed --name "odoo-autocad-integration" --icon "icon/odoo_autocad.ico" --add-data "config;config" --add-data "fonts;fonts" --add-data "icon;icon" --add-data "db;db" --hidden-import "customtkinter" --hidden-import "win32com.client" --hidden-import "win32com.gen_py" --hidden-import "pywintypes" --hidden-import "win32api" --hidden-import "tkinter" --hidden-import "tkinter.ttk" --hidden-import "sqlalchemy" --hidden-import "sqlalchemy.ext.declarative" --hidden-import "sqlalchemy.orm" --exclude-module "pytest" --exclude-module "unittest" --exclude-module "doctest" --exclude-module "pdb" --exclude-module "matplotlib" --exclude-module "numpy" --exclude-module "pandas" --clean --noconfirm --distpath "output" odoo.py
if %ERRORLEVEL% NEQ 0 (
    echo ❌ EXE 建置失敗！
    pause
    exit /b 1
)
REM 複製額外檔案到輸出目錄
copy "doc\ANTIVIRUS_SOLUTION.md" "output\" >nul 2>&1
copy "README.md" "output\" >nul 2>&1
copy "doc\README-DEVELOPMENT.md" "output\" >nul 2>&1
echo ✅ EXE 建置完成
echo.

REM 步驟 2: 程式碼簽章 (如果有憑證)
echo 🔐 步驟 2/3: 程式碼簽章...
if exist "certs\codesign.pfx" (
    echo 正在簽署 output\odoo-autocad-integration.exe...
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe" sign /f "certs\codesign.pfx" /p "YourSecurePassword123!" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "output\odoo-autocad-integration.exe"
    if %ERRORLEVEL% EQU 0 (
        echo ✅ 程式碼簽章完成
    ) else (
        echo ⚠️ 程式碼簽章失敗，但繼續建置
    )
) else (
    echo ℹ️ 未找到憑證檔案 certs\codesign.pfx，跳過簽章
)
echo.

REM 步驟 3: 建立安裝包
echo 📦 步驟 3/3: 建立 Inno Setup 安裝包...
REM 確保使用無簽章版本以避免憑證密碼問題
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "installer\odoo-autocad-setup.iss"
if %ERRORLEVEL% NEQ 0 (
    echo ❌ 安裝包建置失敗！
    pause
    exit /b 1
)
echo ✅ 安裝包建置完成
echo.

REM 顯示完成結果
echo 🎉 建置完成！
echo ==========================================
echo 📁 輸出檔案位置:
echo    EXE: output\odoo-autocad-integration.exe
if exist "installer\odoo-autocad-integration-5.0-setup.exe" (
    echo    安裝包: installer\odoo-autocad-integration-5.0-setup.exe
)
echo.

REM 顯示檔案大小
for %%I in (output\odoo-autocad-integration.exe) do echo 📊 EXE 大小: %%~zI bytes
if exist "installer\odoo-autocad-integration-5.0-setup.exe" (
    for %%I in (installer\odoo-autocad-integration-5.0-setup.exe) do echo 📊 安裝包大小: %%~zI bytes
)
echo.

echo 💡 建議:
echo    1. 測試安裝包在乾淨系統上的安裝
echo    2. 檢查防毒軟體相容性
echo    3. 驗證數位簽章狀態
echo.
pause