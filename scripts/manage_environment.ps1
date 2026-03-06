# Environment Variable Management Script
# Author: GitHub Copilot Assistant
# Date: 2026-02-10

param(
    [string]$Action = "show",  # show, add, remove, clean
    [string]$Path = "",
    [string]$Variable = "PATH",
    [string]$Scope = "User"    # User or Machine
)

function Show-EnvironmentVariable {
    param([string]$VarName, [string]$VarScope)
    
    Write-Host "=== $VarScope 級別的 $VarName 環境變數 ===" -ForegroundColor Green
    $envVar = [Environment]::GetEnvironmentVariable($VarName, [EnvironmentVariableTarget]::$VarScope)
    if ($envVar) {
        $paths = $envVar -split ';'
        for ($i = 0; $i -lt $paths.Count; $i++) {
            if ($paths[$i] -ne "") {
                Write-Host "$($i+1). $($paths[$i])" -ForegroundColor Cyan
            }
        }
    } else {
        Write-Host "環境變數為空" -ForegroundColor Yellow
    }
    Write-Host ""
}

function Add-PathToEnvironment {
    param([string]$NewPath, [string]$VarName, [string]$VarScope)
    
    if (-not (Test-Path $NewPath)) {
        Write-Host "警告：路徑 '$NewPath' 不存在" -ForegroundColor Yellow
        $confirm = Read-Host "是否仍要添加？(y/N)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            return
        }
    }
    
    $currentVar = [Environment]::GetEnvironmentVariable($VarName, [EnvironmentVariableTarget]::$VarScope)
    $paths = $currentVar -split ';' | Where-Object { $_ -ne "" }
    
    if ($paths -contains $NewPath) {
        Write-Host "路徑已存在：$NewPath" -ForegroundColor Yellow
        return
    }
    
    $newVar = ($paths + $NewPath) -join ';'
    [Environment]::SetEnvironmentVariable($VarName, $newVar, [EnvironmentVariableTarget]::$VarScope)
    Write-Host "已添加路徑：$NewPath" -ForegroundColor Green
    
    # 更新當前會話
    $env:PATH = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine) + ";" + [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
}

function Remove-PathFromEnvironment {
    param([string]$PathToRemove, [string]$VarName, [string]$VarScope)
    
    $currentVar = [Environment]::GetEnvironmentVariable($VarName, [EnvironmentVariableTarget]::$VarScope)
    $paths = $currentVar -split ';' | Where-Object { $_ -ne "" -and $_ -ne $PathToRemove }
    
    $newVar = $paths -join ';'
    [Environment]::SetEnvironmentVariable($VarName, $newVar, [EnvironmentVariableTarget]::$VarScope)
    Write-Host "已移除路徑：$PathToRemove" -ForegroundColor Green
    
    # 更新當前會話
    $env:PATH = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine) + ";" + [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
}

function Clean-EnvironmentVariable {
    param([string]$VarName, [string]$VarScope)
    
    $currentVar = [Environment]::GetEnvironmentVariable($VarName, [EnvironmentVariableTarget]::$VarScope)
    $paths = $currentVar -split ';' | Where-Object { 
        $_ -ne "" -and (Test-Path $_)
    } | Sort-Object -Unique
    
    $newVar = $paths -join ';'
    [Environment]::SetEnvironmentVariable($VarName, $newVar, [EnvironmentVariableTarget]::$VarScope)
    Write-Host "已清理重複和無效路徑" -ForegroundColor Green
    
    # 更新當前會話
    $env:PATH = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine) + ";" + [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
}

function Refresh-EnvironmentVariables {
    Write-Host "正在刷新環境變數..." -ForegroundColor Yellow
    
    # 更新當前會話
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine)
    $userPath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
    $env:PATH = $machinePath + ";" + $userPath
    
    # 通知系統環境變數已更改（讓新的應用程式能夠獲取新的環境變數）
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
        $HWND_BROADCAST = [IntPtr]0xffff;
        $WM_SETTINGCHANGE = 0x1a;
        $result = [UIntPtr]::Zero
        [Win32]::SendMessageTimeout($HWND_BROADCAST, $WM_SETTINGCHANGE, [UIntPtr]::Zero, "Environment", 2, 5000, [ref]$result) | Out-Null
        Write-Host "已通知系統環境變數更新" -ForegroundColor Green
    }
    catch {
        Write-Host "注意：無法發送系統通知，新的應用程式可能需要重新啟動才能看到更改" -ForegroundColor Yellow
    }
    
    Write-Host "當前會話環境變數已刷新" -ForegroundColor Green
    Write-Host "提示：新的 terminal 會話可能需要重新啟動才能看到環境變數更改" -ForegroundColor Yellow
}

# 主要邏輯
switch ($Action.ToLower()) {
    "show" {
        Show-EnvironmentVariable -VarName $Variable -VarScope $Scope
        if ($Scope -eq "User") {
            Show-EnvironmentVariable -VarName $Variable -VarScope "Machine"
        }
    }
    "add" {
        if ($Path -eq "") {
            $Path = Read-Host "請輸入要添加的路徑"
        }
        Add-PathToEnvironment -NewPath $Path -VarName $Variable -VarScope $Scope
    }
    "remove" {
        if ($Path -eq "") {
            Show-EnvironmentVariable -VarName $Variable -VarScope $Scope
            $Path = Read-Host "請輸入要移除的路徑"
        }
        Remove-PathFromEnvironment -PathToRemove $Path -VarName $Variable -VarScope $Scope
    }
    "clean" {
        Write-Host "正在清理 $Scope 級別的 $Variable..." -ForegroundColor Yellow
        Clean-EnvironmentVariable -VarName $Variable -VarScope $Scope
    }
    "refresh" {
        Refresh-EnvironmentVariables
    }
    "broadcast" {
        Write-Host "廣播環境變數更新到所有應用程式..." -ForegroundColor Yellow
        Refresh-EnvironmentVariables
    }
    default {
        Write-Host @"
環境變數管理腳本使用說明：

基本用法：
  .\manage_environment.ps1 [參數]

參數：
  -Action     操作類型 (show|add|remove|clean|refresh|broadcast)
  -Path       路徑 (用於 add/remove 操作)
  -Variable   環境變數名稱 (預設：PATH)
  -Scope      範圍 (User|Machine，預設：User)

範例：
  # 顯示當前 PATH
  .\manage_environment.ps1 -Action show

  # 添加路徑到用戶級別 PATH
  .\manage_environment.ps1 -Action add -Path "C:\MyTools"

  # 從系統級別 PATH 移除路徑
  .\manage_environment.ps1 -Action remove -Path "C:\OldTool" -Scope Machine

  # 清理重複和無效路徑
  .\manage_environment.ps1 -Action clean

  # 刷新當前會話的環境變數
  .\manage_environment.ps1 -Action refresh

  # 廣播環境變數更新（通知所有應用程式）
  .\manage_environment.ps1 -Action broadcast
"@ -ForegroundColor Cyan
    }
}
