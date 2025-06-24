# 🏢 企業內部憑證機構 (CA) 建置指南

## 概述
建立企業內部 CA 可以免費簽發程式碼簽章憑證，但僅在企業內部網路中受信任。

## 📋 方案比較

| 方案 | 成本 | 信任範圍 | 複雜度 | 適用場景 |
|------|------|----------|--------|----------|
| **商業 CA** | $200-600/年 | 全球信任 | 低 | 公開發布軟體 |
| **企業 CA** | 免費 | 企業內部 | 中 | 內部部署軟體 |
| **自簽憑證** | 免費 | 單機信任 | 低 | 開發測試 |

## 🏗️ 方案一：Windows Server CA 服務

### 前置需求
- Windows Server 2016/2019/2022
- Active Directory 環境 (建議)
- 企業管理員權限

### 安裝步驟

#### 1. 安裝 AD CS 角色
```powershell
# 使用 PowerShell 安裝
Install-WindowsFeature -Name AD-Certificate -IncludeManagementTools

# 或使用伺服器管理員
# 伺服器管理員 → 新增角色和功能 → Active Directory Certificate Services
```

#### 2. 設定企業根 CA
```powershell
# 安裝企業根 CA
Install-AdcsCertificationAuthority `
    -CAType EnterpriseRootCA `
    -CACommonName "YourCompany Root CA" `
    -KeyLength 4096 `
    -HashAlgorithm SHA256 `
    -CryptoProviderName "RSA#Microsoft Software Key Storage Provider" `
    -ValidityPeriod Years `
    -ValidityPeriodUnits 20
```

#### 3. 設定憑證範本
1. 開啟 **憑證授權單位** 管理工具
2. 展開 CA → 右鍵 **憑證範本** → 管理
3. 右鍵 **程式碼簽章** → 複製範本
4. 設定新範本：
   ```
   範本名稱: Company Code Signing
   主體名稱: 使用 Active Directory 中的資訊
   有效期: 3 年
   金鑰使用方法: 數位簽章、金鑰使用方法的不可否認性
   應用程式原則: 程式碼簽章
   ```

#### 4. 發布憑證範本
```powershell
# 在 CA 上發布範本
Add-CATemplate -Name "Company Code Signing"
```

### 申請程式碼簽章憑證

#### 方法1: 使用 Web 介面
1. 瀏覽器開啟：`http://ca-server/certsrv`
2. 申請憑證 → 進階憑證申請
3. 選擇 "Company Code Signing" 範本
4. 填寫必要資訊並提交

#### 方法2: 使用 MMC
```powershell
# 開啟憑證管理工具
mmc
# 檔案 → 新增/移除嵌入式管理單元 → 憑證 → 電腦帳戶
# 個人 → 憑證 → 右鍵 → 所有工作 → 申請新憑證
```

## 🔧 方案二：OpenSSL 自建 CA

### 安裝 OpenSSL
```bash
# Windows (使用 Chocolatey)
choco install openssl

# 或下載 Win64 OpenSSL
# https://slproweb.com/products/Win32OpenSSL.html
```

### 建立根 CA

#### 1. 設定 CA 目錄結構
```bash
mkdir -p ca/{certs,crl,newcerts,private}
cd ca
echo 1000 > serial
touch index.txt
```

#### 2. 建立 CA 設定檔
```ini
# ca.conf
[req]
default_bits = 4096
prompt = no
distinguished_name = req_distinguished_name
x509_extensions = v3_ca

[req_distinguished_name]
C = TW
ST = Taiwan
L = Taipei
O = Your Company Name
OU = IT Department
CN = Your Company Root CA

[v3_ca]
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid:always,issuer
basicConstraints = critical,CA:true
keyUsage = critical, digitalSignature, cRLSign, keyCertSign

[codesign]
basicConstraints = CA:FALSE
keyUsage = critical, digitalSignature
extendedKeyUsage = critical, codeSigning
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid,issuer
```

#### 3. 產生根 CA 憑證
```bash
# 產生私鑰
openssl genrsa -aes256 -out private/ca.key.pem 4096

# 產生根憑證
openssl req -config ca.conf -key private/ca.key.pem -new -x509 -days 7300 -sha256 -extensions v3_ca -out certs/ca.cert.pem
```

#### 4. 產生程式碼簽章憑證
```bash
# 產生私鑰
openssl genrsa -out private/codesign.key.pem 2048

# 產生憑證申請
openssl req -config ca.conf -key private/codesign.key.pem -new -sha256 -out csr/codesign.csr.pem -subj "/C=TW/ST=Taiwan/L=Taipei/O=Your Company/OU=Development/CN=Code Signing Certificate"

# 簽發憑證
openssl ca -config ca.conf -extensions codesign -days 1095 -notext -md sha256 -in csr/codesign.csr.pem -out certs/codesign.cert.pem
```

#### 5. 轉換為 PFX 格式
```bash
# 合併憑證和私鑰為 PFX
openssl pkcs12 -export -out codesign.pfx -inkey private/codesign.key.pem -in certs/codesign.cert.pem -certfile certs/ca.cert.pem
```

## 🚀 方案三：簡化版自簽憑證

### 快速產生測試憑證
```powershell
# 使用 PowerShell 產生自簽憑證
$cert = New-SelfSignedCertificate `
    -Type CodeSigning `
    -Subject "CN=Your Company Code Signing" `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
    -KeyExportPolicy Exportable `
    -KeyUsage DigitalSignature `
    -ValidityPeriod Years `
    -ValidityPeriodUnits 3 `
    -CertStoreLocation Cert:\CurrentUser\My

# 匯出為 PFX
$password = ConvertTo-SecureString -String "your_password" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "codesign.pfx" -Password $password
```

## 📦 部署企業憑證

### 方法1: 群組原則部署
1. 群組原則管理 → 建立新 GPO
2. 電腦設定 → 原則 → Windows 設定 → 安全性設定 → 公開金鑰原則
3. 受信任的根憑證授權單位 → 匯入根 CA 憑證

### 方法2: 腳本部署
```powershell
# 部署腳本
certlm.msc
Import-Certificate -FilePath "ca.cert.pem" -CertStoreLocation Cert:\LocalMachine\Root
```

### 方法3: 手動安裝
```bash
# 每台電腦手動安裝根憑證
# 雙擊 ca.cert.pem → 安裝憑證 → 本機電腦 → 受信任的根憑證授權單位
```

## 🔐 在 Inno Setup 中使用

### 使用企業憑證簽章
```ini
[Setup]
; 使用企業 CA 簽發的憑證
SignTool=signtool /f "C:\certs\codesign.pfx" /p "password" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f

; 或使用憑證存放區
SignTool=signtool /n "Your Company Code Signing" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f
```

## ⚖️ 方案選擇建議

### 大型企業
```
推薦: Windows Server CA + Active Directory
優點: 集中管理、自動部署、企業級功能
成本: 需要 Windows Server 授權
```

### 中小企業
```
推薦: OpenSSL 自建 CA
優點: 完全免費、跨平台、靈活
成本: 需要技術知識和手動管理
```

### 開發測試
```
推薦: PowerShell 自簽憑證
優點: 快速產生、易於使用
限制: 僅適用於測試環境
```

## 🛡️ 安全考量

### 保護根 CA
- 根 CA 私鑰離線儲存
- 使用 HSM 硬體安全模組
- 定期備份憑證和設定

### 憑證管理
- 設定合理的有效期限
- 建立憑證撤銷清單 (CRL)
- 監控憑證使用情況

### 存取控制
- 限制 CA 管理權限
- 稽核憑證簽發活動
- 實施雙重驗證

## 📝 維護任務

### 定期檢查
- 監控憑證到期日
- 更新 CRL 分發點
- 檢查 CA 健康狀態

### 備份策略
```powershell
# 備份 CA 資料庫
Backup-CARoleService -Path "C:\CABackup"

# 備份憑證
Export-Certificate -Cert $cert -FilePath "backup.cer"
```

## 💡 最佳實務

1. **階層式 CA**: 使用離線根 CA + 線上發行 CA
2. **憑證範本**: 建立標準化的憑證範本
3. **自動化**: 使用 PowerShell 自動化憑證管理
4. **監控**: 設定憑證到期提醒
5. **文件化**: 記錄所有程序和密碼

## 🚨 注意事項

### 企業內部 CA 限制
- 僅在企業內部受信任
- 需要在每台電腦安裝根憑證
- 不適用於公開發布的軟體
- 需要持續維護和管理

### 建議使用場景
- 企業內部應用程式
- 開發和測試環境
- 概念驗證專案
- 預算有限的情況