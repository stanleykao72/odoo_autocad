# Simple PowerShell Script to Create Code Signing Certificates
# Fixed version with proper EKU settings

param(
    [Parameter(Mandatory=$true)]
    [string]$CompanyName,
    
    [Parameter(Mandatory=$true)]
    [string]$CertPassword,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator. Please restart PowerShell as Administrator."
    exit 1
}

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force
}

Write-Host "Creating CA certificates..." -ForegroundColor Green

try {
    # 1. Create Root CA certificate
    Write-Host "Creating Root CA certificate..." -ForegroundColor Yellow
    
    $rootCert = New-SelfSignedCertificate `
        -Subject "CN=Root CA, O=$CompanyName" `
        -KeyUsage CertSign, CRLSign, DigitalSignature `
        -KeyLength 4096 `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears(10) `
        -CertStoreLocation "Cert:\LocalMachine\My"

    # 2. Create Code Signing certificate with proper EKU
    Write-Host "Creating Code Signing certificate..." -ForegroundColor Yellow
    
    $codeCert = New-SelfSignedCertificate `
        -Subject "CN=Code Signing, O=$CompanyName" `
        -Signer $rootCert `
        -KeyUsage DigitalSignature `
        -KeyLength 2048 `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -HashAlgorithm SHA256 `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3") `
        -NotAfter (Get-Date).AddYears(3) `
        -CertStoreLocation "Cert:\LocalMachine\My"

    # 3. Export certificates
    Write-Host "Exporting certificate files..." -ForegroundColor Yellow
    
    # Export Root CA certificate
    $rootCertPath = Join-Path $OutputPath "root-ca.cer"
    Export-Certificate -Cert $rootCert -FilePath $rootCertPath -Type CERT

    # Export Code Signing certificate (public key)
    $codeCertPath = Join-Path $OutputPath "codesign.cer"
    Export-Certificate -Cert $codeCert -FilePath $codeCertPath -Type CERT

    # Export Code Signing certificate (private key PFX)
    $pfxPath = Join-Path $OutputPath "codesign.pfx"
    $securePassword = ConvertTo-SecureString -String $CertPassword -Force -AsPlainText
    Export-PfxCertificate -Cert $codeCert -FilePath $pfxPath -Password $securePassword

    # 4. Create deployment script
    Write-Host "Creating deployment script..." -ForegroundColor Yellow
    
    $deployContent = @"
@echo off
echo Installing Root CA certificate...
certutil -addstore "Root" "root-ca.cer"
if %errorlevel% equ 0 (
    echo Certificate installed successfully!
) else (
    echo Certificate installation failed! Error: %errorlevel%
)
pause
"@
    
    $deployPath = Join-Path $OutputPath "deploy-ca.bat"
    $deployContent | Out-File -FilePath $deployPath -Encoding ASCII

    # 5. Create Inno Setup configuration
    $innoContent = @"
[Setup]
; Code signing configuration
SignTool=signtool sign /f "$pfxPath" /p "$CertPassword" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 `$f
"@
    
    $configPath = Join-Path $OutputPath "inno-setup-config.txt"
    $innoContent | Out-File -FilePath $configPath -Encoding UTF8

    # 6. Create README file
    $readmeContent = @"
CA Certificate Creation Completed!

Files created:
- root-ca.cer: Root CA certificate
- codesign.cer: Code signing certificate (public key)
- codesign.pfx: Code signing certificate (private key)
- deploy-ca.bat: Deployment script
- inno-setup-config.txt: Inno Setup configuration

Usage:
1. Run deploy-ca.bat as administrator
2. Use inno-setup-config.txt in your .iss file
3. Test: signtool sign /f "codesign.pfx" /p "$CertPassword" /fd sha256 "test.exe"

Certificate Thumbprints:
- Root CA: $($rootCert.Thumbprint)
- Code Signing: $($codeCert.Thumbprint)
- Created: $(Get-Date)
"@
    
    $readmePath = Join-Path $OutputPath "README.txt"
    $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8

    # 7. Verify certificate EKU
    Write-Host "Verifying certificate EKU..." -ForegroundColor Yellow
    
    $certDetails = Get-PfxCertificate -FilePath $pfxPath
    $hasCodeSigning = $certDetails.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq "1.3.6.1.5.5.7.3.3" }
    
    if ($hasCodeSigning) {
        Write-Host "Certificate EKU is correct - includes Code Signing" -ForegroundColor Green
    } else {
        Write-Host "WARNING: Certificate may be missing Code Signing EKU" -ForegroundColor Yellow
    }

    Write-Host "`nCA certificates created successfully!" -ForegroundColor Green
    Write-Host "Output directory: $OutputPath" -ForegroundColor Cyan

} catch {
    Write-Error "Error creating certificates: $($_.Exception.Message)"
    exit 1
}

# Optional cleanup
$cleanup = Read-Host "`nRemove certificates from local store? (y/N)"
if ($cleanup -eq 'y' -or $cleanup -eq 'Y') {
    Write-Host "Cleaning up certificate store..." -ForegroundColor Yellow
    Remove-Item -Path "Cert:\LocalMachine\My\$($rootCert.Thumbprint)" -Force
    Remove-Item -Path "Cert:\LocalMachine\My\$($codeCert.Thumbprint)" -Force
    Write-Host "Certificates removed from local store." -ForegroundColor Green
}