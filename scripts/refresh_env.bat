@echo off
echo === 環境變數修復工具 ===
echo.

REM 刷新環境變數
echo 正在刷新環境變數...

REM 重新讀取註冊表中的環境變數
for /f "skip=2 tokens=3*" %%i in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH') do set "MACHINE_PATH=%%i %%j"
for /f "skip=2 tokens=3*" %%i in ('reg query "HKCU\Environment" /v PATH 2^>nul') do set "USER_PATH=%%i %%j"

REM 更新當前會話的PATH
set "PATH=%MACHINE_PATH%;%USER_PATH%"

echo ✅ 環境變數已刷新

echo.
echo === 驗證命令 ===
where dotnet >nul 2>&1
if %errorlevel%==0 (
    echo ✅ dotnet 命令可用
    dotnet --version
) else (
    echo ❌ dotnet 命令不可用
)

where copilot >nul 2>&1
if %errorlevel%==0 (
    echo ✅ copilot 命令可用
    copilot --version
) else (
    echo ❌ copilot 命令不可用
)

echo.
echo === 說明 ===
echo 如果命令仍然不可用，請：
echo 1. 重新啟動 VS Code
echo 2. 或開啟新的命令提示字元/PowerShell
pause
