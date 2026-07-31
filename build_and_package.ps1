# Complete build and package process script v5.0 (PowerShell version)
# One-click build from source code to installer

Write-Host "[BUILD] Odoo-AutoCAD Integration System v6.0 - Complete Build Process" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""

# Check required tools
Write-Host "[INFO] Checking build environment..." -ForegroundColor Cyan

# Check Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "[SUCCESS] Python: $pythonVersion" -ForegroundColor Yellow
} catch {
    Write-Host "[ERROR] Python not installed or not in PATH" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}

# Check PyInstaller
try {
    $pyinstallerVersion = pyinstaller --version 2>&1
    Write-Host "[SUCCESS] PyInstaller: $pyinstallerVersion" -ForegroundColor Yellow
} catch {
    Write-Host "[ERROR] PyInstaller not installed, please run: pip install pyinstaller" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}

# Check Inno Setup
$innoPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $innoPath)) {
    Write-Host "[ERROR] Inno Setup 6 not installed, please install Inno Setup first" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
} else {
    Write-Host "[SUCCESS] Inno Setup 6 installed" -ForegroundColor Yellow
}

# Read version from version.py
$versionLine = Get-Content "version.py" | Select-String 'APP_VERSION\s*=\s*"(.+)"'
if ($versionLine) {
    $appVersion = $versionLine.Matches[0].Groups[1].Value
} else {
    $appVersion = "6.0"
    Write-Host "[WARNING] Could not read version from version.py, using default: $appVersion" -ForegroundColor Yellow
}

Write-Host "[SUCCESS] Build environment check completed" -ForegroundColor Green
Write-Host "[INFO] App version: $appVersion" -ForegroundColor Cyan
Write-Host ""

# Step 1: Build EXE
Write-Host "[BUILD] Step 1/3: Building executable..." -ForegroundColor Cyan
try {
    # Direct PyInstaller build, no dependency on build_windows.ps1
    & python -m PyInstaller --onefile --windowed --name "odoo-autocad-integration" --icon "icon/odoo_autocad.ico" --add-data "fonts;fonts" --add-data "icon;icon" --add-data "libs/autocad-mcp/lisp-code;lisp-code" --paths "libs/autocad-mcp/src" --hidden-import "customtkinter" --hidden-import "win32com.client" --hidden-import "win32com.gen_py" --hidden-import "pywintypes" --hidden-import "win32api" --hidden-import "tkinter" --hidden-import "tkinter.ttk" --hidden-import "sqlalchemy" --hidden-import "sqlalchemy.ext.declarative" --hidden-import "sqlalchemy.orm" --hidden-import "mcp" --hidden-import "mcp.server.fastmcp" --hidden-import "mcp.types" --hidden-import "uvicorn" --hidden-import "starlette" --hidden-import "structlog" --hidden-import "autocad_mcp" --hidden-import "autocad_mcp.server" --hidden-import "autocad_mcp.client" --hidden-import "autocad_mcp.config" --hidden-import "autocad_mcp.backends" --hidden-import "autocad_mcp.backends.file_ipc" --hidden-import "autocad_mcp.backends.ezdxf_backend" --hidden-import "httpx" --hidden-import "anyio" --hidden-import "sniffio" --hidden-import "httpx_sse" --hidden-import "pydantic" --hidden-import "pydantic_settings" --hidden-import "sse_starlette" --exclude-module "pytest" --exclude-module "unittest" --exclude-module "doctest" --exclude-module "pdb" --exclude-module "matplotlib" --exclude-module "numpy" --exclude-module "pandas" --clean --noconfirm --distpath "output" odoo.py
    if ($LASTEXITCODE -eq 0) {
        # Copy additional files to output directory
        Copy-Item "doc\ANTIVIRUS_SOLUTION.md" "output\" -ErrorAction SilentlyContinue
        Copy-Item "README.md" "output\" -ErrorAction SilentlyContinue
        Copy-Item "doc\README-DEVELOPMENT.md" "output\" -ErrorAction SilentlyContinue
        Write-Host "[SUCCESS] EXE build completed" -ForegroundColor Green
    } else {
        throw "PyInstaller returned error code: $LASTEXITCODE"
    }
} catch {
    Write-Host "[ERROR] EXE build failed: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}
Write-Host ""

# Step 2: Code signing (if certificate exists)
Write-Host "[BUILD] Step 2/3: Code signing..." -ForegroundColor Cyan
$certPath = "certs\codesign.pfx"
$signtoolPath = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"

if (Test-Path $certPath) {
    Write-Host "[INFO] Signing output\odoo-autocad-integration.exe..." -ForegroundColor Yellow
    try {
        & $signtoolPath sign /f $certPath /p "YourSecurePassword123!" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "output\odoo-autocad-integration.exe"
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[SUCCESS] Code signing completed" -ForegroundColor Green
        } else {
            Write-Host "[WARNING] Code signing failed, but continuing build" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[WARNING] Signing process error: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "[INFO] Certificate file $certPath not found, skipping signing" -ForegroundColor Gray
}
Write-Host ""

# Step 3: Create installer package
Write-Host "[BUILD] Step 3/3: Creating Inno Setup installer..." -ForegroundColor Cyan
try {
    # Use unsigned version to avoid certificate password issues
    # Pass version from version.py via /D define
    & $innoPath "/DMyAppVersion=$appVersion" "installer\odoo-autocad-setup.iss"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[SUCCESS] Installer build completed" -ForegroundColor Green
    } else {
        throw "Inno Setup returned error code: $LASTEXITCODE"
    }
} catch {
    Write-Host "[ERROR] Installer build failed: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}
Write-Host ""

# Display completion results
Write-Host "[SUCCESS] Build completed!" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host "[INFO] Output file locations:" -ForegroundColor Cyan

# EXE file information
$exePath = "output\odoo-autocad-integration.exe"
if (Test-Path $exePath) {
    $exeInfo = Get-Item $exePath
    Write-Host "   EXE: $exePath" -ForegroundColor White
    Write-Host "   [INFO] EXE size: $($exeInfo.Length.ToString('N0')) bytes" -ForegroundColor White
    
    # Calculate SHA256
    $hash = Get-FileHash -Path $exePath -Algorithm SHA256
    Write-Host "   [INFO] SHA256: $($hash.Hash.Substring(0,16))..." -ForegroundColor White
}

# Installer file information
$installerPath = "installer\odoo-autocad-integration-$appVersion-setup.exe"
if (Test-Path $installerPath) {
    $installerInfo = Get-Item $installerPath
    Write-Host "   Installer: $installerPath" -ForegroundColor White
    Write-Host "   [INFO] Installer size: $($installerInfo.Length.ToString('N0')) bytes" -ForegroundColor White
}

Write-Host ""
# Step 4: Clean up intermediate EXE (installer already contains it)
Write-Host "[BUILD] Step 4: Cleaning up..." -ForegroundColor Cyan
if (Test-Path "output\odoo-autocad-integration.exe") {
    Remove-Item "output\odoo-autocad-integration.exe" -Force
    Write-Host "[SUCCESS] Removed output\odoo-autocad-integration.exe (included in installer)" -ForegroundColor Green
}
Write-Host ""

Write-Host "[TIPS] Recommendations:" -ForegroundColor Yellow
Write-Host "   1. Test installer on clean system" -ForegroundColor White
Write-Host "   2. Check antivirus software compatibility" -ForegroundColor White  
Write-Host "   3. Verify digital signature status" -ForegroundColor White
Write-Host ""

Read-Host "Press Enter to close..."