# Simple CA Certificate Creation Script
# Avoid encoding issues by using English messages and simple syntax

param(
    [string]$CompanyName = "Your Company Name",
    [string]$CertPassword = "SecurePassword123!",
    [string]$OutputPath = ".\certs"
)

Write-Host "Enterprise Internal CA Creation Tool" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: Please run this script as Administrator!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Create output directory
New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
Write-Host "Created output directory: $OutputPath" -ForegroundColor Green

try {
    Write-Host "Creating Root CA certificate..." -ForegroundColor Cyan
    
    # Create Root CA
    $rootCA = New-SelfSignedCertificate `
        -Type Custom `
        -Subject "CN=$CompanyName Root CA, O=$CompanyName, C=TW" `
        -KeyAlgorithm RSA `
        -KeyLength 4096 `
        -KeyExportPolicy Exportable `
        -KeyUsage CertSign, CRLSign, DigitalSignature `
        -ValidityPeriod Years `
        -ValidityPeriodUnits 10 `
        -CertStoreLocation Cert:\LocalMachine\My

    Write-Host "Root CA created successfully: $($rootCA.Thumbprint)" -ForegroundColor Green

    # Install Root CA to Trusted Root
    Write-Host "Installing Root CA to Trusted Root store..." -ForegroundColor Cyan
    $rootCAStore = Get-Item -Path Cert:\LocalMachine\Root
    $rootCAStore.Open("ReadWrite")
    $rootCAStore.Add($rootCA)
    $rootCAStore.Close()
    Write-Host "Root CA installed to Trusted Root store" -ForegroundColor Green

    # Create Code Signing Certificate
    Write-Host "Creating Code Signing certificate..." -ForegroundColor Cyan
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

    Write-Host "Code Signing certificate created: $($codeSignCert.Thumbprint)" -ForegroundColor Green

    # Export certificates
    Write-Host "Exporting certificates..." -ForegroundColor Cyan
    
    $rootCAPath = Join-Path $OutputPath "root-ca.cer"
    Export-Certificate -Cert $rootCA -FilePath $rootCAPath | Out-Null
    Write-Host "Root CA exported: $rootCAPath" -ForegroundColor White

    $codeSignCertPath = Join-Path $OutputPath "codesign.cer"
    Export-Certificate -Cert $codeSignCert -FilePath $codeSignCertPath | Out-Null
    Write-Host "Code Signing cert exported: $codeSignCertPath" -ForegroundColor White

    # Export PFX
    $pfxPath = Join-Path $OutputPath "codesign.pfx"
    $securePassword = ConvertTo-SecureString -String $CertPassword -Force -AsPlainText
    Export-PfxCertificate -Cert $codeSignCert -FilePath $pfxPath -Password $securePassword | Out-Null
    Write-Host "PFX file exported: $pfxPath" -ForegroundColor White

    # Create deployment script
    $deployScriptPath = Join-Path $OutputPath "deploy-ca.bat"
    @"
@echo off
echo Installing Enterprise Root CA certificate...
powershell -Command "Import-Certificate -FilePath 'root-ca.cer' -CertStoreLocation Cert:\LocalMachine\Root"
if %ERRORLEVEL% EQU 0 (
    echo Root CA certificate installed successfully
) else (
    echo Root CA certificate installation failed
)
pause
"@ | Out-File -FilePath $deployScriptPath -Encoding ASCII
    Write-Host "Deployment script created: $deployScriptPath" -ForegroundColor White

    # Create Inno Setup config
    $innoConfigPath = Join-Path $OutputPath "inno-setup-config.txt"
    $pfxFullPath = (Resolve-Path $pfxPath).Path
    @"
; Enterprise Internal Certificate Inno Setup Configuration
[Setup]
; ... other settings ...

; Code signing configuration
SignTool=signtool /f "$pfxFullPath" /p "$CertPassword" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f

; Alternative: Using certificate store (when certificate is installed)
; SignTool=signtool /n "$CompanyName Code Signing" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f
"@ | Out-File -FilePath $innoConfigPath -Encoding UTF8
    Write-Host "Inno Setup config created: $innoConfigPath" -ForegroundColor White

    # Create README
    $readmePath = Join-Path $OutputPath "README.txt"
    @"
Enterprise Internal CA Certificates
====================================

Files Generated:
- root-ca.cer: Root CA certificate (deploy to all computers)
- codesign.cer: Code signing certificate (public key)
- codesign.pfx: Code signing certificate (with private key for signing)
- deploy-ca.bat: Deployment script for Root CA
- inno-setup-config.txt: Inno Setup signing configuration

Usage Steps:

1. Deploy Root CA Certificate
   Run deploy-ca.bat on each computer that needs to trust this certificate

2. Configure Inno Setup
   Add the content from inno-setup-config.txt to your .iss file

3. Test Code Signing
   signtool sign /f "codesign.pfx" /p "$CertPassword" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "your-file.exe"

Certificate Information:
- Company Name: $CompanyName
- PFX Password: $CertPassword
- Root CA Thumbprint: $($rootCA.Thumbprint)
- Code Signing Thumbprint: $($codeSignCert.Thumbprint)

Security Notes:
- Keep PFX file and password secure
- Backup certificates regularly
- Monitor certificate usage
- Renew before expiration
"@ | Out-File -FilePath $readmePath -Encoding UTF8
    Write-Host "README created: $readmePath" -ForegroundColor White

    # Display summary
    Write-Host ""
    Write-Host "SUCCESS: Enterprise CA setup completed!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Output directory: $OutputPath" -ForegroundColor Cyan
    Write-Host "Root CA thumbprint: $($rootCA.Thumbprint)" -ForegroundColor Yellow
    Write-Host "Code signing thumbprint: $($codeSignCert.Thumbprint)" -ForegroundColor Yellow
    Write-Host "PFX password: $CertPassword" -ForegroundColor Red
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor White
    Write-Host "1. Run deploy-ca.bat on other computers to install Root CA" -ForegroundColor White
    Write-Host "2. Use codesign.pfx for signing your software" -ForegroundColor White
    Write-Host "3. Refer to inno-setup-config.txt for Inno Setup configuration" -ForegroundColor White

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Press Enter to exit..." -ForegroundColor Gray
Read-Host