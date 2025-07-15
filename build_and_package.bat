@echo off
REM Complete build and package process script v5.0
REM One-click build from source code to installer

echo [BUILD] Odoo-AutoCAD Integration System v5.0 - Complete Build Process
echo ===============================================
echo.

REM Check required tools
echo [INFO] Checking build environment...

echo [LOG] Checking Python...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python not installed or not in PATH
    pause
    exit /b 1
)
echo [LOG] Python check passed

echo [LOG] Checking PyInstaller...
pyinstaller --version
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller not installed, please run: pip install pyinstaller
    pause
    exit /b 1
)
echo [LOG] PyInstaller check passed

echo [LOG] Checking Inno Setup...
if not exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo [ERROR] Inno Setup 6 not installed, please install Inno Setup first
    pause
    exit /b 1
)
echo [LOG] Inno Setup check passed

echo [SUCCESS] Build environment check completed
echo.

REM Step 1: Build EXE
echo [BUILD] Step 1/3: Building executable...
echo [LOG] Starting PyInstaller...
REM Direct PyInstaller build, no dependency on build_windows.bat
python -m PyInstaller --onefile --windowed --name "odoo-autocad-integration" --icon "icon/odoo_autocad.ico" --add-data "config;config" --add-data "fonts;fonts" --add-data "icon;icon" --add-data "db;db" --hidden-import "customtkinter" --hidden-import "win32com.client" --hidden-import "win32com.gen_py" --hidden-import "pywintypes" --hidden-import "win32api" --hidden-import "tkinter" --hidden-import "tkinter.ttk" --hidden-import "sqlalchemy" --hidden-import "sqlalchemy.ext.declarative" --hidden-import "sqlalchemy.orm" --exclude-module "pytest" --exclude-module "unittest" --exclude-module "doctest" --exclude-module "pdb" --exclude-module "matplotlib" --exclude-module "numpy" --exclude-module "pandas" --clean --noconfirm --distpath "output" odoo.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] EXE build failed!
    pause
    exit /b 1
)
echo [LOG] PyInstaller build completed
echo [LOG] Copying additional files to output directory...
REM Copy additional files to output directory
copy "doc\ANTIVIRUS_SOLUTION.md" "output\" >nul 2>&1
copy "README.md" "output\" >nul 2>&1
copy "doc\README-DEVELOPMENT.md" "output\" >nul 2>&1
echo [LOG] File copy completed
echo [SUCCESS] EXE build completed
echo.

REM Step 2: Code signing (if certificate exists)
echo [BUILD] Step 2/3: Code signing...
echo [LOG] Checking certificate file...
if exist "certs\codesign.pfx" (
    echo [LOG] Certificate file found, starting signing...
    echo [INFO] Signing output\odoo-autocad-integration.exe...
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe" sign /f "certs\codesign.pfx" /p "YourSecurePassword123!" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "output\odoo-autocad-integration.exe"
    if %ERRORLEVEL% EQU 0 (
        echo [SUCCESS] Code signing completed
    ) else (
        echo [WARNING] Code signing failed, but continuing build
    )
) else (
    echo [INFO] Certificate file certs\codesign.pfx not found, skipping signing
)
echo [LOG] Signing step completed
echo.

REM Step 3: Create installer package
echo [BUILD] Step 3/3: Creating Inno Setup installer...
echo [LOG] Starting Inno Setup...
REM Use unsigned version to avoid certificate password issues
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "installer\odoo-autocad-setup.iss"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Installer build failed!
    pause
    exit /b 1
)
echo [LOG] Inno Setup build completed
echo [SUCCESS] Installer build completed
echo.

REM Display completion results
echo [SUCCESS] Build completed!
echo ==========================================
echo [INFO] Output file locations:
echo    EXE: output\odoo-autocad-integration.exe
if exist "installer\odoo-autocad-integration-5.0-setup.exe" (
    echo    Installer: installer\odoo-autocad-integration-5.0-setup.exe
)
echo.

REM Display file sizes
for %%I in (output\odoo-autocad-integration.exe) do echo [INFO] EXE size: %%~zI bytes
if exist "installer\odoo-autocad-integration-5.0-setup.exe" (
    for %%I in (installer\odoo-autocad-integration-5.0-setup.exe) do echo [INFO] Installer size: %%~zI bytes
)
echo.

echo [TIPS] Recommendations:
echo    1. Test installer on clean system
echo    2. Check antivirus software compatibility
echo    3. Verify digital signature status
echo.
pause