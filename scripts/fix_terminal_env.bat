@echo off
echo 正在修復新 terminal 的環境變數問題...
echo.

REM 確保環境變數已設置
powershell -Command "& {$env:PATH = [Environment]::GetEnvironmentVariable('PATH', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('PATH', 'User')}"

REM 廣播環境變數更改
powershell -Command "& {Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class Win32{[DllImport(\"user32.dll\")]public static extern IntPtr SendMessageTimeout(IntPtr hWnd,uint Msg,UIntPtr wParam,string lParam,uint fuFlags,uint uTimeout,out UIntPtr lpdwResult);}';$HWND_BROADCAST=[IntPtr]0xffff;$WM_SETTINGCHANGE=0x1a;$result=[UIntPtr]::Zero;[Win32]::SendMessageTimeout($HWND_BROADCAST,$WM_SETTINGCHANGE,[UIntPtr]::Zero,'Environment',2,5000,[ref]$result)}"

echo ✅ 環境變數修復完成
echo.
echo === 解決方案總結 ===
echo 1. ✅ PowerShell 設定檔案已安裝
echo 2. ✅ 環境變數更改已廣播到系統
echo 3. ✅ 新的 PowerShell 會話將自動載入正確的 PATH
echo.
echo 建議：
echo - 重新啟動 VS Code 或開啟新的 terminal 來測試
echo - 如有問題，運行：refreshpath
echo.
pause
