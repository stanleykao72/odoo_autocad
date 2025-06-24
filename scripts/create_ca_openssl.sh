#!/bin/bash
# 使用 OpenSSL 建立企業內部 CA
# 適用於 Linux/macOS/Windows (WSL/Git Bash)

set -e

# 配置變數
COMPANY_NAME="Your Company Name"
COUNTRY="TW"
STATE="Taiwan"
CITY="Taipei"
ORG_UNIT="IT Department"
CA_DAYS=3650  # 10年
CERT_DAYS=1095  # 3年
KEY_SIZE=4096
CERT_KEY_SIZE=2048

# 顏色定義
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo -e "${GREEN}🏢 企業內部 CA 建置工具 (OpenSSL)${NC}"
echo -e "${GREEN}======================================${NC}"

# 檢查 OpenSSL
if ! command -v openssl &> /dev/null; then
    echo -e "${RED}❌ OpenSSL 未安裝或不在 PATH 中${NC}"
    echo -e "${YELLOW}請安裝 OpenSSL：${NC}"
    echo -e "  Windows: choco install openssl"
    echo -e "  macOS:   brew install openssl"
    echo -e "  Ubuntu:  sudo apt-get install openssl"
    exit 1
fi

echo -e "${GREEN}✅ OpenSSL 版本: $(openssl version)${NC}"

# 建立目錄結構
echo -e "${CYAN}📁 建立目錄結構...${NC}"
mkdir -p ca/{certs,crl,newcerts,private,csr}
cd ca
echo 1000 > serial
touch index.txt
echo -e "${GREEN}✅ 目錄結構建立完成${NC}"

# 建立 CA 設定檔
echo -e "${CYAN}📝 建立 CA 設定檔...${NC}"
cat > ca.conf << EOF
[req]
default_bits = $KEY_SIZE
prompt = no
distinguished_name = req_distinguished_name
x509_extensions = v3_ca

[req_distinguished_name]
C = $COUNTRY
ST = $STATE
L = $CITY
O = $COMPANY_NAME
OU = $ORG_UNIT
CN = $COMPANY_NAME Root CA

[ca]
default_ca = CA_default

[CA_default]
dir = .
certs = \$dir/certs
crl_dir = \$dir/crl
database = \$dir/index.txt
new_certs_dir = \$dir/newcerts
certificate = \$dir/certs/ca.cert.pem
serial = \$dir/serial
crlnumber = \$dir/crlnumber
crl = \$dir/crl/ca.crl.pem
private_key = \$dir/private/ca.key.pem
RANDFILE = \$dir/private/.rand
x509_extensions = usr_cert
name_opt = ca_default
cert_opt = ca_default
default_days = $CERT_DAYS
default_crl_days = 30
default_md = sha256
preserve = no
policy = policy_match

[policy_match]
countryName = match
stateOrProvinceName = match
organizationName = match
organizationalUnitName = optional
commonName = supplied
emailAddress = optional

[usr_cert]
basicConstraints = CA:FALSE
nsComment = "OpenSSL Generated Certificate"
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid,issuer

[v3_ca]
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid:always,issuer
basicConstraints = critical,CA:true
keyUsage = critical, digitalSignature, cRLSign, keyCertSign

[codesign_req]
default_bits = $CERT_KEY_SIZE
prompt = no
distinguished_name = codesign_distinguished_name
req_extensions = v3_req

[codesign_distinguished_name]
C = $COUNTRY
ST = $STATE
L = $CITY
O = $COMPANY_NAME
OU = Development
CN = $COMPANY_NAME Code Signing

[v3_req]
basicConstraints = CA:FALSE
keyUsage = nonRepudiation, digitalSignature, keyEncipherment

[codesign]
basicConstraints = CA:FALSE
keyUsage = critical, digitalSignature
extendedKeyUsage = critical, codeSigning
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid,issuer
nsComment = "Code Signing Certificate"
EOF

echo -e "${GREEN}✅ CA 設定檔建立完成${NC}"

# 產生根 CA 私鑰
echo -e "${CYAN}🔐 產生根 CA 私鑰...${NC}"
openssl genrsa -aes256 -out private/ca.key.pem $KEY_SIZE
echo -e "${GREEN}✅ 根 CA 私鑰產生完成${NC}"

# 產生根 CA 憑證
echo -e "${CYAN}📜 產生根 CA 憑證...${NC}"
openssl req -config ca.conf -key private/ca.key.pem -new -x509 -days $CA_DAYS -sha256 -extensions v3_ca -out certs/ca.cert.pem
echo -e "${GREEN}✅ 根 CA 憑證產生完成${NC}"

# 驗證根 CA 憑證
echo -e "${CYAN}🔍 驗證根 CA 憑證...${NC}"
openssl x509 -noout -text -in certs/ca.cert.pem

# 產生程式碼簽章憑證私鑰
echo -e "${CYAN}🔑 產生程式碼簽章憑證私鑰...${NC}"
openssl genrsa -out private/codesign.key.pem $CERT_KEY_SIZE
echo -e "${GREEN}✅ 程式碼簽章憑證私鑰產生完成${NC}"

# 產生程式碼簽章憑證申請
echo -e "${CYAN}📋 產生程式碼簽章憑證申請...${NC}"
openssl req -config ca.conf -section codesign_req -key private/codesign.key.pem -new -sha256 -out csr/codesign.csr.pem
echo -e "${GREEN}✅ 憑證申請產生完成${NC}"

# 簽發程式碼簽章憑證
echo -e "${CYAN}✍️ 簽發程式碼簽章憑證...${NC}"
openssl ca -config ca.conf -extensions codesign -days $CERT_DAYS -notext -md sha256 -in csr/codesign.csr.pem -out certs/codesign.cert.pem
echo -e "${GREEN}✅ 程式碼簽章憑證簽發完成${NC}"

# 產生 PFX 檔案
echo -e "${CYAN}📦 產生 PFX 檔案...${NC}"
openssl pkcs12 -export -out codesign.pfx -inkey private/codesign.key.pem -in certs/codesign.cert.pem -certfile certs/ca.cert.pem
echo -e "${GREEN}✅ PFX 檔案產生完成${NC}"

# 建立部署腳本
echo -e "${CYAN}📜 建立部署腳本...${NC}"

# Windows 部署腳本
cat > deploy-ca-windows.bat << 'EOF'
@echo off
echo 正在安裝企業根 CA 憑證到 Windows...

REM 安裝到當前使用者的受信任根憑證授權單位
certutil -user -addstore "Root" ca.cert.pem
if %ERRORLEVEL% EQU 0 (
    echo ✅ 根 CA 憑證安裝成功 (使用者)
) else (
    echo ❌ 根 CA 憑證安裝失敗 (使用者)
)

REM 安裝到本機的受信任根憑證授權單位 (需要管理員權限)
certutil -addstore "Root" ca.cert.pem
if %ERRORLEVEL% EQU 0 (
    echo ✅ 根 CA 憑證安裝成功 (本機)
) else (
    echo ⚠️ 根 CA 憑證安裝失敗 (本機) - 可能需要管理員權限
)

echo.
echo 完成！您現在可以使用企業憑證簽章軟體。
pause
EOF

# Linux/macOS 部署腳本
cat > deploy-ca-linux.sh << 'EOF'
#!/bin/bash
echo "正在安裝企業根 CA 憑證到系統..."

# 複製憑證到系統目錄
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    # Linux
    sudo cp certs/ca.cert.pem /usr/local/share/ca-certificates/company-root-ca.crt
    sudo update-ca-certificates
    echo "✅ 根 CA 憑證已安裝到 Linux 系統"
elif [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS
    sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain certs/ca.cert.pem
    echo "✅ 根 CA 憑證已安裝到 macOS 系統"
else
    echo "❌ 不支援的作業系統"
    exit 1
fi

echo "完成！您現在可以使用企業憑證簽章軟體。"
EOF

chmod +x deploy-ca-linux.sh

# 建立 README
cat > README.md << EOF
# 企業內部 CA 憑證

## 檔案說明

### 憑證檔案
- \`certs/ca.cert.pem\`: 根 CA 憑證 (需要部署到所有電腦)
- \`certs/codesign.cert.pem\`: 程式碼簽章憑證 (公鑰)
- \`codesign.pfx\`: 程式碼簽章憑證 (含私鑰，用於簽章)

### 私鑰檔案 (請妥善保管)
- \`private/ca.key.pem\`: 根 CA 私鑰
- \`private/codesign.key.pem\`: 程式碼簽章私鑰

### 部署檔案
- \`deploy-ca-windows.bat\`: Windows 根 CA 憑證部署腳本
- \`deploy-ca-linux.sh\`: Linux/macOS 根 CA 憑證部署腳本

## 使用步驟

### 1. 部署根 CA 憑證

#### Windows
\`\`\`
deploy-ca-windows.bat
\`\`\`

#### Linux/macOS
\`\`\`
./deploy-ca-linux.sh
\`\`\`

### 2. 簽章檔案
\`\`\`bash
# 使用 OpenSSL 簽章 (較複雜)
openssl dgst -sha256 -sign private/codesign.key.pem -out signature.sig your-file.exe

# 使用 signtool (Windows，需要先安裝 Windows SDK)
signtool sign /f codesign.pfx /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 your-file.exe
\`\`\`

### 3. 在 Inno Setup 中使用
\`\`\`ini
[Setup]
SignTool=signtool /f "codesign.pfx" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 \$f
\`\`\`

## 憑證資訊
- 公司名稱: $COMPANY_NAME
- 根 CA 有效期: $CA_DAYS 天
- 簽章憑證有效期: $CERT_DAYS 天

## 安全注意事項
- 妥善保管 private/ 目錄中的私鑰檔案
- 定期備份整個 CA 目錄
- 監控憑證使用情況
- 憑證到期前及時更新

## 驗證憑證
\`\`\`bash
# 檢視根 CA 憑證
openssl x509 -noout -text -in certs/ca.cert.pem

# 檢視程式碼簽章憑證
openssl x509 -noout -text -in certs/codesign.cert.pem

# 驗證憑證鏈
openssl verify -CAfile certs/ca.cert.pem certs/codesign.cert.pem
\`\`\`
EOF

echo -e "${GREEN}✅ 部署腳本和說明文件建立完成${NC}"

# 顯示摘要
echo -e "\n${GREEN}🎉 企業 CA 建置完成！${NC}"
echo -e "${GREEN}======================================${NC}"
echo -e "${CYAN}輸出目錄: $(pwd)${NC}"
echo -e "${YELLOW}根 CA 憑證: certs/ca.cert.pem${NC}"
echo -e "${YELLOW}程式碼簽章憑證: certs/codesign.cert.pem${NC}"
echo -e "${RED}PFX 檔案: codesign.pfx${NC}"

echo -e "\n${BLUE}📋 下一步操作：${NC}"
echo -e "1. 在其他電腦執行部署腳本安裝根 CA"
echo -e "2. 使用 codesign.pfx 簽章您的軟體"
echo -e "3. 參考 README.md 了解詳細使用方法"

echo -e "\n${YELLOW}⚠️  重要提醒：${NC}"
echo -e "- 請妥善保管 private/ 目錄中的私鑰"
echo -e "- 建議將整個 CA 目錄備份到安全位置"
echo -e "- PFX 檔案密碼請記錄在安全地方"