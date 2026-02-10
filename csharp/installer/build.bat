@echo off
REM ===================================================================
REM Odoo AutoCAD Integration (C#) - Build and Package
REM Usage:  build.bat [options]
REM Options:
REM   --installer-only   Skip dotnet publish, use existing publish dir
REM   --skip-tests       Skip running tests
REM   --keep-symbols     Keep debug symbols in output
REM   --clean            Clean bin/obj before build (slower but safer)
REM ===================================================================

setlocal enabledelayedexpansion

echo.
echo ========================================
echo  Odoo AutoCAD C# Build Script
echo ========================================
echo.

REM --- Parse Arguments ---
set INSTALLER_ONLY=0
set SKIP_TESTS=0
set KEEP_SYMBOLS=0
set DO_CLEAN=0

:parse_args
if "%~1"=="" goto done_args
if /i "%~1"=="--installer-only" set INSTALLER_ONLY=1
if /i "%~1"=="--skip-tests" set SKIP_TESTS=1
if /i "%~1"=="--keep-symbols" set KEEP_SYMBOLS=1
if /i "%~1"=="--clean" set DO_CLEAN=1
shift
goto parse_args
:done_args

REM --- Configuration ---
set SCRIPT_DIR=%~dp0
set PROJECT_ROOT=%SCRIPT_DIR%..\OdooAutoCADIntegration
set APP_PROJECT=%PROJECT_ROOT%\src\OdooAutoCAD.App\OdooAutoCAD.App.csproj
set TEST_PROJECT=%PROJECT_ROOT%\tests\OdooAutoCAD.Integration.Tests\OdooAutoCAD.Integration.Tests.csproj
set PUBLISH_DIR=%SCRIPT_DIR%..\publish
set INSTALLER_SCRIPT=%SCRIPT_DIR%odoo-autocad-csharp-setup.iss
set INNO_SETUP="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

REM --- Locate dotnet ---
set DOTNET=dotnet
where dotnet >nul 2>&1
if errorlevel 1 (
    if exist "C:\Program Files\dotnet\dotnet.exe" (
        set DOTNET="C:\Program Files\dotnet\dotnet.exe"
    ) else (
        echo ERROR: .NET SDK not found.
        echo   Install from: https://dotnet.microsoft.com/download/dotnet/8.0
        exit /b 1
    )
)

REM --- Installer-Only Mode ---
if %INSTALLER_ONLY%==1 (
    echo Mode: Installer only ^(skip build^)
    echo.
    if not exist "%PUBLISH_DIR%\OdooAutoCAD.exe" (
        echo ERROR: publish\OdooAutoCAD.exe not found.
        echo   Run without --installer-only first.
        exit /b 1
    )
    goto step_installer
)

REM --- Step 1: Check Prerequisites ---
echo [1/5] Checking prerequisites...

for /f "tokens=*" %%v in ('%DOTNET% --version 2^>nul') do set DOTNET_VER=%%v
echo    .NET SDK: %DOTNET_VER%

if not exist %INNO_SETUP% (
    echo ERROR: Inno Setup not found.
    echo   Install from: https://jrsoftware.org/isdl.php
    exit /b 1
)
echo    Inno Setup: OK
echo.

REM --- Step 2: Clean (Optional) ---
if %DO_CLEAN%==1 (
    echo [2/5] Cleaning previous build...
    if exist "%PUBLISH_DIR%" rmdir /s /q "%PUBLISH_DIR%"
    for /d /r "%PROJECT_ROOT%\src" %%d in (bin obj) do (
        if exist "%%d" rmdir /s /q "%%d" 2>nul
    )
    echo    Done
) else (
    echo [2/5] Skipping clean ^(use --clean to force^)
)
echo.

REM --- Step 3: Restore ---
echo [3/5] Restoring NuGet packages...
%DOTNET% restore "%PROJECT_ROOT%\OdooAutoCADIntegration.sln" --verbosity quiet
if errorlevel 1 (
    echo ERROR: NuGet restore failed
    exit /b 1
)
echo    Done
echo.

REM --- Step 3.5: Tests (Optional) ---
if %SKIP_TESTS%==0 (
    if exist "%TEST_PROJECT%" (
        echo [3.5] Running tests...
        %DOTNET% test "%TEST_PROJECT%" --configuration Release --no-restore --verbosity normal
        if errorlevel 1 (
            echo WARNING: Tests failed. Continuing build...
        ) else (
            echo    All tests passed
        )
        echo.
    )
)

REM --- Step 4: Publish ---
echo [4/5] Publishing application...
echo    Config:  Release / win-x64 / Self-Contained
echo    Output:  %PUBLISH_DIR%

set PUBLISH_ARGS=--configuration Release --runtime win-x64 --self-contained true --output "%PUBLISH_DIR%" /p:PublishReadyToRun=true /p:IncludeNativeLibrariesForSelfExtract=true

if %KEEP_SYMBOLS%==0 (
    set PUBLISH_ARGS=%PUBLISH_ARGS% /p:DebugType=none /p:DebugSymbols=false
)

%DOTNET% publish "%APP_PROJECT%" %PUBLISH_ARGS%
if errorlevel 1 (
    echo ERROR: Build failed
    exit /b 1
)
echo    Done
echo.

REM --- Step 5: Installer ---
:step_installer
echo [5/5] Creating installer...

%INNO_SETUP% "%INSTALLER_SCRIPT%"
if errorlevel 1 (
    echo ERROR: Installer creation failed
    exit /b 1
)
echo.

REM --- Summary ---
echo ========================================
echo  BUILD SUCCESSFUL
echo ========================================

REM Read version from .iss file
for /f "tokens=2 delims= " %%v in ('findstr /c:"#define MyAppVersion" "%INSTALLER_SCRIPT%"') do set APP_VER=%%~v

set INSTALLER_EXE=%SCRIPT_DIR%odoo-autocad-integration-%APP_VER%-setup.exe

if exist "%INSTALLER_EXE%" (
    for %%F in ("%INSTALLER_EXE%") do (
        set /a SIZE_MB=%%~zF / 1048576
        echo  Installer: %%~nxF
        echo  Size:      !SIZE_MB! MB
        echo  Path:      %%~fF
    )
)
echo ========================================
echo.

endlocal
