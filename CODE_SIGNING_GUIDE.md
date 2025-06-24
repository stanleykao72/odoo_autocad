# 🔐 程式碼簽章配置指南

## 概述
程式碼簽章可以大幅減少防毒軟體誤報，增加用戶對軟體的信任度。

## 📋 前置準備

### 1. 安裝 Windows SDK
下載並安裝 Windows SDK 以獲得 `signtool.exe`：
```
https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
```

### 2. 獲取程式碼簽章憑證
- **商業憑證**：從 DigiCert、GlobalSign、Sectigo 等購買
- **企業憑證**：使用公司內部 CA 簽發
- **測試憑證**：自簽憑證（僅供開發測試）

## 🛠️ SignTool 配置方法

### 方法1: 使用 PFX 檔案
```ini
[Setup]
SignTool=signtool /f "C:\path\to\certificate.pfx" /p "password" /t "http://timestamp.digicert.com" $f
```

**參數說明：**
- `/f "path"`: PFX 憑證檔案路徑
- `/p "password"`: PFX 檔案密碼
- `/t "url"`: 時間戳記伺服器（重要！）
- `$f`: Inno Setup 會自動替換為要簽章的檔案

### 方法2: 使用憑證存放區
```ini
[Setup]
SignTool=signtool /n "Your Company Name" /t "http://timestamp.digicert.com" $f
```

**參數說明：**
- `/n "name"`: 憑證主體名稱
- 憑證需要先安裝到 Windows 憑證存放區

### 方法3: 使用 SHA256 (推薦)
```ini
[Setup]
SignTool=signtool /sha1 "thumbprint" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f
```

**參數說明：**
- `/sha1 "thumbprint"`: 憑證指紋
- `/fd sha256`: 使用 SHA256 摘要算法
- `/tr "url"`: RFC 3161 時間戳記伺服器
- `/td sha256`: 時間戳記摘要算法

## 📝 實際配置範例

### 範例1: 使用 PFX 檔案
```ini
; 在 odoo-autocad-setup.iss 中
[Setup]
; ... 其他設定 ...
SignTool=signtool /f "C:\certs\company.pfx" /p "MySecretPassword" /t "http://timestamp.digicert.com" $f
```

### 範例2: 使用環境變數保護密碼
```ini
; 設定環境變數
; set CERT_PASSWORD=MySecretPassword

[Setup]
SignTool=signtool /f "C:\certs\company.pfx" /p "%CERT_PASSWORD%" /t "http://timestamp.digicert.com" $f
```

### 範例3: 同時簽章安裝程式和 EXE
```ini
[Setup]
; 簽章安裝程式
SignTool=signtool /f "C:\certs\company.pfx" /p "password" /t "http://timestamp.digicert.com" $f

[Files]
; 先簽章 EXE 檔案，再打包
Source: "C:\odoo\autocad_source\output\odoo-autocad-integration.exe"; DestDir: "{app}"; Flags: ignoreversion
```

## 🕐 時間戳記伺服器

### 常用的時間戳記伺服器：
```
DigiCert:     http://timestamp.digicert.com
GlobalSign:   http://timestamp.globalsign.com/scripts/timstamp.dll
Sectigo:      http://timestamp.sectigo.com
Microsoft:    http://timestamp.microsoft.com/
```

### RFC 3161 時間戳記伺服器（推薦）：
```
DigiCert:     http://timestamp.digicert.com
GlobalSign:   http://rfc3161timestamp.globalsign.com/advanced
Sectigo:      http://timestamp.sectigo.com/rfc3161
```

## 🔧 SignTool 完整參數

### 基本參數：
```bash
signtool sign [options] <file(s)>

# 常用選項：
/f <file>           # PFX 憑證檔案
/p <password>       # PFX 密碼
/n <name>           # 憑證主體名稱
/sha1 <hash>        # 憑證指紋
/fd <algorithm>     # 檔案摘要算法 (sha1, sha256)
/t <url>            # 時間戳記伺服器 (舊格式)
/tr <url>           # RFC 3161 時間戳記伺服器
/td <algorithm>     # 時間戳記摘要算法
/d <description>    # 簽章描述
/du <url>           # 簽章 URL
```

## 📋 完整設定步驟

### 步驟1: 準備憑證
```bash
# 如果使用 PFX 檔案，確保路徑正確
dir "C:\path\to\certificate.pfx"

# 測試 signtool 可用性
signtool /?
```

### 步驟2: 更新 ISS 檔案
```ini
[Setup]
; 其他設定...
SignTool=signtool /f "C:\certs\company.pfx" /p "password" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f
```

### 步驟3: 編譯安裝程式
```bash
# 使用 Inno Setup Compiler
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "config\odoo-autocad-setup.iss"
```

## 🛡️ 安全最佳實務

### 1. 密碼保護
```ini
; 使用環境變數
SignTool=signtool /f "C:\certs\company.pfx" /p "%CERT_PASSWORD%" /tr "http://timestamp.digicert.com" $f

; 或使用憑證存放區（無需密碼）
SignTool=signtool /n "Company Name" /tr "http://timestamp.digicert.com" $f
```

### 2. 憑證存放
- 將 PFX 檔案放在安全位置
- 設定適當的檔案權限
- 考慮使用 HSM (硬體安全模組)

### 3. 建置腳本
```bash
# 建置腳本範例
@echo off
echo 正在建置 EXE...
pyinstaller --onefile odoo.py

echo 正在簽章 EXE...
signtool sign /f "C:\certs\company.pfx" /p "%CERT_PASSWORD%" /tr "http://timestamp.digicert.com" "dist\odoo.exe"

echo 正在建立安裝程式...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "config\odoo-autocad-setup.iss"

echo 完成！
```

## 🧪 測試簽章

### 驗證簽章：
```bash
# 檢查檔案簽章
signtool verify /pa "path\to\signed\file.exe"

# 詳細資訊
signtool verify /v /pa "path\to\signed\file.exe"
```

### 檢查憑證：
```bash
# 列出憑證存放區中的憑證
certlm.msc
```

## ❌ 常見錯誤

### 1. "SignTool Error: No certificates were found"
- 檢查憑證路徑或名稱
- 確認憑證已安裝到正確的存放區

### 2. "SignTool Error: The specified timestamp server"
- 更換時間戳記伺服器
- 檢查網路連接

### 3. "Access denied"
- 檢查檔案權限
- 以系統管理員身分執行

## 💰 憑證成本參考

| 供應商 | 價格範圍 | 有效期 |
|--------|----------|--------|
| DigiCert | $400-600/年 | 1-3年 |
| GlobalSign | $300-500/年 | 1-3年 |
| Sectigo | $200-400/年 | 1-3年 |
| 企業 CA | 免費 | 依政策 |