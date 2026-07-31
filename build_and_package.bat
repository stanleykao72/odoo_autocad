@echo off
REM Complete build and package process script v5.0
REM One-click build from source code to installer

echo [BUILD] Odoo-AutoCAD Integration System v6.0 - Complete Build Process
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

REM Read version from version.py
for /f "tokens=2 delims== " %%A in ('findstr "APP_VERSION" version.py') do set APP_VERSION=%%~A
if "%APP_VERSION%"=="" set APP_VERSION=6.0
echo [INFO] App version: %APP_VERSION%
echo [SUCCESS] Build environment check completed
echo.

REM Step 1: Build EXE
echo [BUILD] Step 1/3: Building executable...
echo [LOG] Starting PyInstaller...
REM Remove any stale EXE first: otherwise a failed build would leave the previous
REM binary in place and the later steps would sign and package that old file.
if exist "output\odoo-autocad-integration.exe" del /f /q "output\odoo-autocad-integration.exe"
if exist "output\odoo-autocad-integration.exe" (
    echo [ERROR] Cannot delete stale output\odoo-autocad-integration.exe - is it running?
    pause
    exit /b 1
)
REM All packaging settings live in odoo-autocad-integration.spec
python -m PyInstaller --clean --noconfirm --distpath "output" "odoo-autocad-integration.spec"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] EXE build failed!
    pause
    exit /b 1
)
if not exist "output\odoo-autocad-integration.exe" (
    echo [ERROR] PyInstaller reported success but output\odoo-autocad-integration.exe is missing
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
REM
REM Notes on the block below - do NOT move these comments inside the parentheses:
REM   * A REM line containing quote characters inside a parenthesised block
REM     breaks cmd's parsing of the whole if/else chain.
REM   * Use "if errorlevel 1" rather than "if %ERRORLEVEL% EQU 0": the latter is
REM     expanded when the block is parsed, so it always sees a stale value and
REM     would report signing success even when signtool failed.
REM   * The PFX password comes from the CODESIGN_PASSWORD environment variable.
REM     Never hardcode it here - this file is tracked in git.
echo [BUILD] Step 2/3: Code signing...
echo [LOG] Checking certificate file...
if not exist "certs\codesign.pfx" (
    echo [INFO] Certificate file certs\codesign.pfx not found, skipping signing
) else if "%CODESIGN_PASSWORD%"=="" (
    REM Never hardcode the password here - this file is tracked in git.
    echo [WARNING] CODESIGN_PASSWORD not set, skipping signing
    echo [INFO] This session only:  set CODESIGN_PASSWORD=your_pfx_password
    echo [INFO] Permanent:          setx CODESIGN_PASSWORD "your_pfx_password"
) else (
    echo [LOG] Certificate file found, starting signing...
    echo [INFO] Signing output\odoo-autocad-integration.exe...
    "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe" sign /f "certs\codesign.pfx" /p "%CODESIGN_PASSWORD%" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "output\odoo-autocad-integration.exe"
    if errorlevel 1 (
        echo [WARNING] Code signing failed, but continuing build
    ) else (
        echo [SUCCESS] Code signing completed
    )
)
echo [LOG] Signing step completed
echo.

REM Step 3: Create installer package
echo [BUILD] Step 3/3: Creating Inno Setup installer...
echo [LOG] Starting Inno Setup...
REM Use unsigned version to avoid certificate password issues
REM Pass version from version.py via /D define
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=%APP_VERSION% "installer\odoo-autocad-setup.iss"
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
if exist "installer\odoo-autocad-integration-%APP_VERSION%-setup.exe" (
    echo    Installer: installer\odoo-autocad-integration-%APP_VERSION%-setup.exe
)
echo.

REM Display file sizes
for %%I in (output\odoo-autocad-integration.exe) do echo [INFO] EXE size: %%~zI bytes
if exist "installer\odoo-autocad-integration-%APP_VERSION%-setup.exe" (
    for %%I in (installer\odoo-autocad-integration-%APP_VERSION%-setup.exe) do echo [INFO] Installer size: %%~zI bytes
)
echo.

REM Step 4: Clean up intermediate EXE
echo [BUILD] Step 4: Cleaning up...
if exist "output\odoo-autocad-integration.exe" (
    del /f "output\odoo-autocad-integration.exe"
    echo [SUCCESS] Removed output\odoo-autocad-integration.exe (included in installer)
)
echo.

echo [TIPS] Recommendations:
echo    1. Test installer on clean system
echo    2. Check antivirus software compatibility
echo    3. Verify digital signature status
echo.
pause