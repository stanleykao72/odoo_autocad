# Refresh Environment Variables for New Sessions
# This script fixes the issue where new terminal sessions don't see updated environment variables

Write-Host "=== 環境變數修復工具 ===" -ForegroundColor Green

# Function to broadcast environment variable changes
function Send-EnvironmentUpdateBroadcast {
    try {
        Add-Type -TypeDefinition @'
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern IntPtr SendMessageTimeout(
                IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
                uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
        }
'@
        $HWND_BROADCAST = [IntPtr]0xffff
        $WM_SETTINGCHANGE = 0x1a
        $result = [UIntPtr]::Zero
        [Win32]::SendMessageTimeout($HWND_BROADCAST, $WM_SETTINGCHANGE, [UIntPtr]::Zero, "Environment", 2, 5000, [ref]$result) | Out-Null
        Write-Host "✅ 環境變數更新已廣播到系統" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "⚠️ 無法廣播環境變數更新: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# Get current environment variables
$machinePath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine)
$userPath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)

Write-Host "檢查重要路徑..." -ForegroundColor Yellow

# Check for important paths
$importantPaths = @(
    "C:\Program Files\dotnet",
    "C:\Users\$env:USERNAME\AppData\Local\Microsoft\WinGet\Links"
)

$pathsToAdd = @()
$allPaths = ($machinePath + ";" + $userPath) -split ';' | Where-Object { $_ -ne "" }

foreach ($importantPath in $importantPaths) {
    $expandedPath = [Environment]::ExpandEnvironmentVariables($importantPath)
    if ($allPaths -notcontains $expandedPath) {
        if (Test-Path $expandedPath) {
            $pathsToAdd += $expandedPath
            Write-Host "❌ 缺少路徑: $expandedPath" -ForegroundColor Red
        } else {
            Write-Host "⚠️ 路徑不存在: $expandedPath" -ForegroundColor Yellow
        }
    } else {
        Write-Host "✅ 路徑存在: $expandedPath" -ForegroundColor Green
    }
}

# Add missing paths to user PATH
if ($pathsToAdd.Count -gt 0) {
    Write-Host "正在添加缺少的路徑..." -ForegroundColor Yellow
    $currentUserPath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
    $userPaths = if ($currentUserPath) { $currentUserPath -split ';' | Where-Object { $_ -ne "" } } else { @() }
    
    foreach ($pathToAdd in $pathsToAdd) {
        if ($userPaths -notcontains $pathToAdd) {
            $userPaths += $pathToAdd
            Write-Host "➕ 添加到用戶 PATH: $pathToAdd" -ForegroundColor Cyan
        }
    }
    
    $newUserPath = $userPaths -join ';'
    [Environment]::SetEnvironmentVariable("PATH", $newUserPath, [EnvironmentVariableTarget]::User)
    Write-Host "✅ 用戶 PATH 已更新" -ForegroundColor Green
}

# Update current session
$env:PATH = $machinePath + ";" + [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
Write-Host "✅ 當前會話 PATH 已更新" -ForegroundColor Green

# Broadcast changes
Write-Host "廣播環境變數更改..." -ForegroundColor Yellow
if (Send-EnvironmentUpdateBroadcast) {
    Write-Host "✅ 環境變數更改已廣播" -ForegroundColor Green
} else {
    Write-Host "⚠️ 請重新啟動 VS Code 或新的 terminal 以獲取更新的環境變數" -ForegroundColor Yellow
}

Write-Host "`n=== 驗證命令可用性 ===" -ForegroundColor Green

# Test commands
$commands = @("dotnet", "copilot")
foreach ($cmd in $commands) {
    $cmdPath = Get-Command $cmd -ErrorAction SilentlyContinue
    if ($cmdPath) {
        Write-Host "✅ $cmd 可用: $($cmdPath.Source)" -ForegroundColor Green
    } else {
        Write-Host "❌ $cmd 不可用" -ForegroundColor Red
    }
}

Write-Host "`n=== 建議 ===" -ForegroundColor Cyan
Write-Host "1. 如果命令仍然不可用，請重新啟動 VS Code" -ForegroundColor White
Write-Host "2. 或者開啟新的 PowerShell 窗口" -ForegroundColor White
Write-Host "3. 運行此腳本: .\scripts\fix_env.ps1" -ForegroundColor White
