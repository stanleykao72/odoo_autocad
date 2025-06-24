@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM 企業內部 CA 建置腳本 - Batch 版本
REM 使用 PowerShell 命令建立程式碼簽章憑證

echo 🏢 企業內部 CA 建置工具
echo ================================

REM 檢查參數
set "COMPANY_NAME=%~1"
set "CERT_PASSWORD=%~2"
set "OUTPUT_PATH=%~3"

if "%COMPANY_NAME%"=="" set "COMPANY_NAME=Your Company Name"
if "%CERT_PASSWORD%"=="" set "CERT_PASSWORD=SecurePassword123!"
if "%OUTPUT_PATH%"=="" set "OUTPUT_PATH=.\certs"

echo 公司名稱: %COMPANY_NAME%
echo 憑證密碼: %CERT_PASSWORD%
echo 輸出目錄: %OUTPUT_PATH%
echo.

REM 檢查是否以管理員身分執行
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 請以管理員身分執行此腳本！
    pause
    exit /b 1
)

REM 建立輸出目錄
if not exist "%OUTPUT_PATH%" mkdir "%OUTPUT_PATH%"
echo ✅ 建立輸出目錄: %OUTPUT_PATH%

echo.
echo 🔐 建立根 CA 憑證...

REM 建立 PowerShell 腳本內容
set "PS_SCRIPT=%TEMP%\create_ca_temp.ps1"

(
echo # 建立根 CA 憑證
echo $rootCA = New-SelfSignedCertificate `
echo     -Type Custom `
echo     -Subject "CN=%COMPANY_NAME% Root CA, O=%COMPANY_NAME%, C=TW" `
echo     -KeyAlgorithm RSA `
echo     -KeyLength 4096 `
echo     -KeyExportPolicy Exportable `
echo     -KeyUsage CertSign, CRLSign, DigitalSignature `
echo     -ValidityPeriod Years `
echo     -ValidityPeriodUnits 10 `
echo     -CertStoreLocation Cert:\LocalMachine\My
echo.
echo Write-Host "根 CA 憑證建立成功: $($rootCA.Thumbprint^)"
echo.
echo # 安裝根 CA 到受信任的根憑證授權單位
echo $rootCAStore = Get-Item -Path Cert:\LocalMachine\Root
echo $rootCAStore.Open("ReadWrite"^)
echo $rootCAStore.Add($rootCA^)
echo $rootCAStore.Close(^)
echo Write-Host "根 CA 已安裝到受信任的根憑證授權單位"
echo.
echo # 建立程式碼簽章憑證
echo $codeSignCert = New-SelfSignedCertificate `
echo     -Type CodeSigningCert `
echo     -Subject "CN=%COMPANY_NAME% Code Signing, O=%COMPANY_NAME%, C=TW" `
echo     -KeyAlgorithm RSA `
echo     -KeyLength 2048 `
echo     -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
echo     -KeyExportPolicy Exportable `
echo     -KeyUsage DigitalSignature `
echo     -ValidityPeriod Years `
echo     -ValidityPeriodUnits 3 `
echo     -CertStoreLocation Cert:\LocalMachine\My `
echo     -Signer $rootCA
echo.
echo Write-Host "程式碼簽章憑證建立成功: $($codeSignCert.Thumbprint^)"
echo.
echo # 匯出憑證
echo $rootCAPath = Join-Path "%OUTPUT_PATH%" "root-ca.cer"
echo Export-Certificate -Cert $rootCA -FilePath $rootCAPath ^| Out-Null
echo Write-Host "根 CA 憑證已匯出: $rootCAPath"
echo.
echo $codeSignCertPath = Join-Path "%OUTPUT_PATH%" "codesign.cer"
echo Export-Certificate -Cert $codeSignCert -FilePath $codeSignCertPath ^| Out-Null
echo Write-Host "程式碼簽章憑證已匯出: $codeSignCertPath"
echo.
echo # 匯出 PFX 檔案
echo $pfxPath = Join-Path "%OUTPUT_PATH%" "codesign.pfx"
echo $securePassword = ConvertTo-SecureString -String "%CERT_PASSWORD%" -Force -AsPlainText
echo Export-PfxCertificate -Cert $codeSignCert -FilePath $pfxPath -Password $securePassword ^| Out-Null
echo Write-Host "PFX 檔案已匯出: $pfxPath"
echo.
echo # 建立部署腳本
echo $deployScript = @'
echo @echo off
echo echo 正在安裝企業根 CA 憑證...
echo powershell -Command "Import-Certificate -FilePath 'root-ca.cer' -CertStoreLocation Cert:\LocalMachine\Root"
echo if %%ERRORLEVEL%% EQU 0 (
echo     echo 根 CA 憑證安裝成功
echo ^) else (
echo     echo 根 CA 憑證安裝失敗
echo ^)
echo pause
echo '@
echo.
echo $deployScriptPath = Join-Path "%OUTPUT_PATH%" "deploy-ca.bat"
echo $deployScript ^| Out-File -FilePath $deployScriptPath -Encoding ASCII
echo Write-Host "部署腳本已建立: $deployScriptPath"
echo.
echo # 建立 Inno Setup 配置範例
echo $innoConfig = @"
echo ; 使用企業內部憑證的 Inno Setup 配置
echo [Setup]
echo ; ... 其他設定 ...
echo.
echo ; 程式碼簽章配置
echo SignTool=signtool /f "$pfxPath" /p "%CERT_PASSWORD%" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f
echo.
echo ; 或使用憑證存放區 (憑證已安裝時^)
echo ; SignTool=signtool /n "%COMPANY_NAME% Code Signing" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f
echo "@
echo.
echo $innoConfigPath = Join-Path "%OUTPUT_PATH%" "inno-setup-config.txt"
echo $innoConfig ^| Out-File -FilePath $innoConfigPath -Encoding UTF8
echo Write-Host "Inno Setup 配置已建立: $innoConfigPath"
echo.
echo # 顯示摘要
echo Write-Host ""
echo Write-Host "🎉 企業 CA 建置完成！"
echo Write-Host "================================"
echo Write-Host "輸出目錄: %OUTPUT_PATH%"
echo Write-Host "根 CA 指紋: $($rootCA.Thumbprint^)"
echo Write-Host "簽章憑證指紋: $($codeSignCert.Thumbprint^)"
echo Write-Host "PFX 密碼: %CERT_PASSWORD%"
echo Write-Host ""
echo Write-Host "📋 下一步操作："
echo Write-Host "1. 在其他電腦執行 deploy-ca.bat 安裝根 CA"
echo Write-Host "2. 使用 codesign.pfx 簽章您的軟體"
echo Write-Host "3. 參考 inno-setup-config.txt 配置 Inno Setup"
) > "%PS_SCRIPT%"

echo 📝 執行 PowerShell 憑證建立程序...
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

if %errorlevel% neq 0 (
    echo ❌ 憑證建立失敗
    del "%PS_SCRIPT%" 2>nul
    pause
    exit /b 1
)

echo.
echo ✅ 企業 CA 建置完成！

REM 清理暫存檔
del "%PS_SCRIPT%" 2>nul

echo.
echo 📋 產生的檔案：
echo   - %OUTPUT_PATH%\root-ca.cer (根 CA 憑證)
echo   - %OUTPUT_PATH%\codesign.cer (程式碼簽章憑證)
echo   - %OUTPUT_PATH%\codesign.pfx (PFX 檔案，含私鑰)
echo   - %OUTPUT_PATH%\deploy-ca.bat (部署腳本)
echo   - %OUTPUT_PATH%\inno-setup-config.txt (Inno Setup 配置)

echo.
echo 按任意鍵結束...
pause >nul