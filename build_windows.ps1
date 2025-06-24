# Odoo-AutoCAD Integration Build Script for Windows PowerShell
# 減少防毒軟體誤報的優化建置腳本

Write-Host "🚀 開始建置 Odoo-AutoCAD 整合工具..." -ForegroundColor Green
Write-Host ""

# 清理舊的建置檔案
$dirsToClean = @("build", "dist", "output")
foreach ($dir in $dirsToClean) {
    if (Test-Path $dir) {
        Remove-Item -Recurse -Force $dir
        Write-Host "✅ 清理目錄: $dir" -ForegroundColor Yellow
    }
}

# 清理 spec 檔案
Get-ChildItem -Filter "*.spec" | Remove-Item -Force

# 建立輸出目錄
New-Item -ItemType Directory -Path "output" -Force | Out-Null

Write-Host "✅ 清理完成" -ForegroundColor Green

# PyInstaller 參數
$pyinstallerArgs = @(
    "--onefile"
    "--windowed"
    "--name", "odoo-autocad-integration"
    "--icon", "icon/odoo_autocad.ico"
    "--add-data", "config;config"
    "--add-data", "fonts;fonts"
    "--add-data", "icon;icon"
    "--hidden-import", "customtkinter"
    "--hidden-import", "win32com.client"
    "--hidden-import", "win32com.gen_py"
    "--hidden-import", "pywintypes"
    "--hidden-import", "win32api"
    "--hidden-import", "tkinter"
    "--hidden-import", "tkinter.ttk"
    "--exclude-module", "pytest"
    "--exclude-module", "unittest"
    "--exclude-module", "doctest"
    "--exclude-module", "pdb"
    "--exclude-module", "matplotlib"
    "--exclude-module", "numpy"
    "--exclude-module", "pandas"
    "--clean"
    "--noconfirm"
    "--distpath", "output"
    "odoo.py"
)

# 執行 PyInstaller
Write-Host "🔨 執行 PyInstaller..." -ForegroundColor Cyan
try {
    & pyinstaller @pyinstallerArgs
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ PyInstaller 建置成功" -ForegroundColor Green
    } else {
        throw "PyInstaller 回傳錯誤代碼: $LASTEXITCODE"
    }
} catch {
    Write-Host "❌ 建置失敗: $_" -ForegroundColor Red
    Read-Host "按任意鍵繼續..."
    exit 1
}

# 複製額外檔案
Write-Host "📁 複製輸出檔案..." -ForegroundColor Cyan
$filesToCopy = @("ANTIVIRUS_SOLUTION.md")
foreach ($file in $filesToCopy) {
    if (Test-Path $file) {
        Copy-Item $file "output\"
    }
}

# 複製所有 README 檔案
Get-ChildItem -Filter "README*.md" | Copy-Item -Destination "output\"

# 顯示建置結果
Write-Host ""
Write-Host "🎉 建置完成！" -ForegroundColor Green
Write-Host ""

$exePath = "output\odoo-autocad-integration.exe"
if (Test-Path $exePath) {
    $fileInfo = Get-Item $exePath
    Write-Host "📁 檔案位置: $exePath" -ForegroundColor Cyan
    Write-Host "📊 檔案大小: $($fileInfo.Length.ToString('N0')) bytes" -ForegroundColor Cyan
    
    # 計算 SHA256 雜湊
    $hash = Get-FileHash -Path $exePath -Algorithm SHA256
    Write-Host "🔒 SHA256: $($hash.Hash)" -ForegroundColor Cyan
    
    # 建立建置資訊檔案
    $buildInfo = @{
        version = "3.0"
        build_date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        file_size = $fileInfo.Length
        sha256_hash = $hash.Hash
        filename = $fileInfo.Name
        build_environment = @{
            powershell_version = $PSVersionTable.PSVersion.ToString()
            platform = [System.Environment]::OSVersion.ToString()
            architecture = [System.Environment]::Is64BitOperatingSystem ? "64-bit" : "32-bit"
        }
    }
    
    $buildInfo | ConvertTo-Json -Depth 3 | Out-File -FilePath "output\build_info.json" -Encoding UTF8
    Write-Host "📝 建置資訊: output\build_info.json" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "💡 減少防毒軟體誤報建議:" -ForegroundColor Yellow
Write-Host "   1. 將 output 資料夾加入防毒軟體排除清單" -ForegroundColor White
Write-Host "   2. 參考 ANTIVIRUS_SOLUTION.md 設定指南" -ForegroundColor White  
Write-Host "   3. 考慮使用程式碼簽章憑證" -ForegroundColor White
Write-Host ""

Read-Host "按 Enter 鍵關閉..."