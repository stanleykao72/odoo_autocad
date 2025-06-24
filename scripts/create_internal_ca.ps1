# 企業內部 CA 建置腳本
# 使用 PowerShell 快速建立程式碼簽章憑證

param(
    [string]$CompanyName = "Your Company Name",
    [string]$CertPassword = "SecurePassword123!",
    [string]$OutputPath = ".\certs"
)

Write-Host "🏢 企業內部 CA 建置工具" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green

# 檢查是否以管理員身分執行
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "❌ 請以管理員身分執行此腳本！" -ForegroundColor Red
    exit 1
}

# 建立輸出目錄
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
Write-Host "✅ 建立輸出目錄: $OutputPath" -ForegroundColor Green

try {
    # 1. 建立根 CA 憑證
    Write-Host "🔐 建立根 CA 憑證..." -ForegroundColor Cyan
    $rootCA = New-SelfSignedCertificate `
        -Type Custom `
        -Subject "CN=$CompanyName Root CA, O=$CompanyName, C=TW" `
        -KeyAlgorithm RSA `
        -KeyLength 4096 `
        -KeyExportPolicy Exportable `
        -KeyUsage CertSign, CRLSign, DigitalSignature `
        -ValidityPeriod Years `
        -ValidityPeriodUnits 10 `
        -CertStoreLocation Cert:\LocalMachine\My `
        -Extension @(
            New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($true, $true, 0, $true),
            New-Object System.Security.Cryptography.X509Certificates.X509KeyUsageExtension([System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::CertSign -bor [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::CRLSign, $true)
        )

    Write-Host "✅ 根 CA 憑證建立成功: $($rootCA.Thumbprint)" -ForegroundColor Green

    # 2. 將根 CA 安裝到受信任的根憑證授權單位
    Write-Host "🔧 安裝根 CA 到受信任的根憑證授權單位..." -ForegroundColor Cyan
    $rootCAStore = Get-Item -Path Cert:\LocalMachine\Root
    $rootCAStore.Open("ReadWrite")
    $rootCAStore.Add($rootCA)
    $rootCAStore.Close()

    # 3. 建立程式碼簽章憑證
    Write-Host "📝 建立程式碼簽章憑證..." -ForegroundColor Cyan
    $codeSignCert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject "CN=$CompanyName Code Signing, O=$CompanyName, C=TW" `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
        -KeyExportPolicy Exportable `
        -KeyUsage DigitalSignature `
        -ValidityPeriod Years `
        -ValidityPeriodUnits 3 `
        -CertStoreLocation Cert:\LocalMachine\My `
        -Signer $rootCA

    Write-Host "✅ 程式碼簽章憑證建立成功: $($codeSignCert.Thumbprint)" -ForegroundColor Green

    # 4. 匯出憑證
    Write-Host "💾 匯出憑證檔案..." -ForegroundColor Cyan
    
    # 匯出根 CA 憑證
    $rootCAPath = Join-Path $OutputPath "root-ca.cer"
    Export-Certificate -Cert $rootCA -FilePath $rootCAPath | Out-Null
    Write-Host "   根 CA 憑證: $rootCAPath" -ForegroundColor White

    # 匯出程式碼簽章憑證 (公鑰)
    $codeSignCertPath = Join-Path $OutputPath "codesign.cer"
    Export-Certificate -Cert $codeSignCert -FilePath $codeSignCertPath | Out-Null
    Write-Host "   程式碼簽章憑證: $codeSignCertPath" -ForegroundColor White

    # 匯出程式碼簽章憑證 (PFX 含私鑰)
    $pfxPath = Join-Path $OutputPath "codesign.pfx"
    $securePassword = ConvertTo-SecureString -String $CertPassword -Force -AsPlainText
    Export-PfxCertificate -Cert $codeSignCert -FilePath $pfxPath -Password $securePassword | Out-Null
    Write-Host "   PFX 檔案: $pfxPath" -ForegroundColor White

    # 5. 建立部署腳本
    Write-Host "📜 建立部署腳本..." -ForegroundColor Cyan
    $deployScript = @"
@echo off
REM 企業根 CA 憑證部署腳本
echo 正在安裝企業根 CA 憑證...

REM 安裝根 CA 憑證到受信任的根憑證授權單位
certlm.msc
powershell -Command "Import-Certificate -FilePath 'root-ca.cer' -CertStoreLocation Cert:\LocalMachine\Root"

if %ERRORLEVEL% EQU 0 (
    echo ✅ 根 CA 憑證安裝成功
) else (
    echo ❌ 根 CA 憑證安裝失敗
)

pause
"@

    $deployScriptPath = Join-Path $OutputPath "deploy-ca.bat"
    $deployScript | Out-File -FilePath $deployScriptPath -Encoding ASCII
    Write-Host "   部署腳本: $deployScriptPath" -ForegroundColor White

    # 6. 建立 Inno Setup 配置範例
    $innoSetupConfig = @"
; 使用企業內部憑證的 Inno Setup 配置
[Setup]
; ... 其他設定 ...

; 程式碼簽章配置
SignTool=signtool /f "$pfxPath" /p "$CertPassword" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f

; 或使用憑證存放區 (憑證已安裝時)
; SignTool=signtool /n "$CompanyName Code Signing" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f
"@

    $innoConfigPath = Join-Path $OutputPath "inno-setup-config.txt"
    $innoSetupConfig | Out-File -FilePath $innoConfigPath -Encoding UTF8
    Write-Host "   Inno Setup 配置: $innoConfigPath" -ForegroundColor White

    # 7. 建立說明文件
    $readme = @"
# 企業內部 CA 憑證

## 檔案說明

### 憑證檔案
- `root-ca.cer`: 根 CA 憑證 (需要部署到所有電腦)
- `codesign.cer`: 程式碼簽章憑證 (公鑰)
- `codesign.pfx`: 程式碼簽章憑證 (含私鑰，用於簽章)

### 部署檔案
- `deploy-ca.bat`: 根 CA 憑證部署腳本
- `inno-setup-config.txt`: Inno Setup 簽章配置範例

## 使用步驟

### 1. 部署根 CA 憑證
在每台需要信任此憑證的電腦上執行：
```
deploy-ca.bat
```

### 2. 配置 Inno Setup
將 `inno-setup-config.txt` 中的內容加入您的 .iss 檔案

### 3. 簽章檔案
使用以下命令測試簽章：
```
signtool sign /f "codesign.pfx" /p "$CertPassword" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "your-file.exe"
```

## 憑證資訊
- 公司名稱: $CompanyName
- PFX 密碼: $CertPassword
- 根 CA 指紋: $($rootCA.Thumbprint)
- 簽章憑證指紋: $($codeSignCert.Thumbprint)

## 安全注意事項
- 妥善保管 PFX 檔案和密碼
- 定期備份憑證
- 監控憑證使用情況
- 憑證到期前及時更新
"@

    $readmePath = Join-Path $OutputPath "README.md"
    $readme | Out-File -FilePath $readmePath -Encoding UTF8
    Write-Host "   說明文件: $readmePath" -ForegroundColor White

    # 顯示摘要
    Write-Host "`n🎉 企業 CA 建置完成！" -ForegroundColor Green
    Write-Host "================================" -ForegroundColor Green
    Write-Host "輸出目錄: $OutputPath" -ForegroundColor Cyan
    Write-Host "根 CA 指紋: $($rootCA.Thumbprint)" -ForegroundColor Yellow
    Write-Host "簽章憑證指紋: $($codeSignCert.Thumbprint)" -ForegroundColor Yellow
    Write-Host "PFX 密碼: $CertPassword" -ForegroundColor Red
    Write-Host "`n📋 下一步操作：" -ForegroundColor White
    Write-Host "1. 在其他電腦執行 deploy-ca.bat 安裝根 CA" -ForegroundColor White
    Write-Host "2. 使用 codesign.pfx 簽章您的軟體" -ForegroundColor White
    Write-Host "3. 參考 inno-setup-config.txt 配置 Inno Setup" -ForegroundColor White

} catch {
    Write-Host "❌ 發生錯誤: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "`n按任意鍵結束..." -ForegroundColor Gray
Read-Host