@echo off
REM Odoo-AutoCAD Integration Build Script for Windows
REM 減少防毒軟體誤報的優化建置腳本

echo 🚀 開始建置 Odoo-AutoCAD 整合工具...
echo.

REM 清理舊的建置檔案
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist output rmdir /s /q output
if exist *.spec del *.spec

echo ✅ 清理完成

REM 建立輸出目錄
mkdir output

REM 執行 PyInstaller 建置
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "odoo-autocad-integration" ^
    --icon "icon/odoo_autocad.ico" ^
    --add-data "config;config" ^
    --add-data "fonts;fonts" ^
    --add-data "icon;icon" ^
    --add-data "db;db" ^
    --hidden-import "customtkinter" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32com.gen_py" ^
    --hidden-import "pywintypes" ^
    --hidden-import "win32api" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "sqlalchemy" ^
    --hidden-import "sqlalchemy.ext.declarative" ^
    --hidden-import "sqlalchemy.orm" ^
    --exclude-module "pytest" ^
    --exclude-module "unittest" ^
    --exclude-module "doctest" ^
    --exclude-module "pdb" ^
    --exclude-module "matplotlib" ^
    --exclude-module "numpy" ^
    --exclude-module "pandas" ^
    --clean ^
    --noconfirm ^
    --distpath "output" ^
    odoo.py

if %ERRORLEVEL% NEQ 0 (
    echo ❌ 建置失敗！
    pause
    exit /b 1
)

echo.
echo ✅ 建置成功！

REM 複製額外檔案到輸出目錄
copy "ANTIVIRUS_SOLUTION.md" "output\"
copy "README*.md" "output\" 2>nul

REM 顯示檔案資訊
echo.
echo 📁 檔案位置: output\odoo-autocad-integration.exe
echo 📊 檔案大小:
for %%I in (output\odoo-autocad-integration.exe) do echo    %%~zI bytes

echo.
echo 🎉 建置完成！
echo.
echo 💡 減少防毒軟體誤報建議:
echo    1. 將 output 資料夾加入防毒軟體排除清單
echo    2. 參考 ANTIVIRUS_SOLUTION.md 設定指南
echo    3. 考慮使用程式碼簽章憑證
echo.
pause