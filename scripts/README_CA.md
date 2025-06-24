# 企業 CA 憑證建立工具

## 使用方法

### Batch 腳本 (推薦)
```cmd
REM 以管理員身分執行 Command Prompt，然後執行：
scripts\create_internal_ca.bat "承暉精品股份有限公司" "YourSecurePassword123!" ".\certs"
```

### PowerShell 腳本 (如果有編碼問題請使用 Batch 版本)
```powershell
# 以管理員身分執行 PowerShell，然後執行：
.\scripts\create_internal_ca.ps1 -CompanyName "承暉精品股份有限公司" -CertPassword "YourSecurePassword123!" -OutputPath ".\certs"
```

## 參數說明

1. **公司名稱**: 會出現在憑證的主體名稱中
2. **憑證密碼**: PFX 檔案的密碼，用於保護私鑰
3. **輸出目錄**: 憑證檔案的存放目錄

## 產生的檔案

執行後會在指定目錄產生以下檔案：

- `root-ca.cer` - 根 CA 憑證 (需要部署到所有電腦)
- `codesign.cer` - 程式碼簽章憑證 (公鑰)
- `codesign.pfx` - 程式碼簽章憑證 (含私鑰，用於簽章)
- `deploy-ca.bat` - 部署腳本，用於在其他電腦安裝根 CA
- `inno-setup-config.txt` - Inno Setup 簽章配置範例

## 使用步驟

### 1. 建立憑證
```cmd
scripts\create_internal_ca.bat "承暉精品股份有限公司" "YourSecurePassword123!" ".\certs"
```

### 2. 部署根 CA 到其他電腦
將 `root-ca.cer` 和 `deploy-ca.bat` 複製到需要信任此憑證的電腦，然後以管理員身分執行：
```cmd
deploy-ca.bat
```

### 3. 配置 Inno Setup
將 `inno-setup-config.txt` 中的內容加入您的 `.iss` 檔案：
```ini
[Setup]
SignTool=signtool /f "C:\path\to\codesign.pfx" /p "YourSecurePassword123!" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 $f
```

### 4. 測試簽章
```cmd
signtool sign /f "certs\codesign.pfx" /p "YourSecurePassword123!" /fd sha256 /tr "http://timestamp.digicert.com" /td sha256 "your-file.exe"
```

## 注意事項

- 必須以**管理員身分**執行腳本
- 妥善保管 PFX 檔案和密碼
- 在企業內部所有電腦安裝根 CA 憑證
- 定期備份憑證檔案
- 憑證到期前及時更新

## 故障排除

### 1. "拒絕存取" 錯誤
- 確保以管理員身分執行腳本

### 2. PowerShell 執行原則錯誤
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. 中文字符顯示問題
- 使用 Batch 版本 (create_internal_ca.bat)
- 確保 Command Prompt 使用 UTF-8 編碼

### 4. 憑證不受信任
- 確保在所有需要的電腦上執行 `deploy-ca.bat`
- 檢查憑證是否正確安裝到 "受信任的根憑證授權單位"