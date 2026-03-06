@echo off
REM build.bat — Build odoo_bridge.exe using PyInstaller
REM Output: ../dist/odoo_bridge.exe

echo ========================================
echo  Building odoo_bridge.exe
echo ========================================

cd /d "%~dp0"

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH
    exit /b 1
)

REM Check PyInstaller
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM Install requirements
echo Installing requirements...
pip install -r requirements.txt

REM Build
echo Building executable...
python -m PyInstaller ^
    --onefile ^
    --name odoo_bridge ^
    --clean ^
    --noconfirm ^
    --distpath "../dist" ^
    --workpath "./build" ^
    --specpath "./build" ^
    --hidden-import configparser ^
    --hidden-import requests ^
    odoo_bridge.py

if errorlevel 1 (
    echo ERROR: Build failed
    exit /b 1
)

echo ========================================
echo  Build complete: ../dist/odoo_bridge.exe
echo ========================================

REM Cleanup
if exist build rmdir /s /q build

pause
