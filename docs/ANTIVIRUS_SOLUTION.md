# 🛡️ 防毒軟體誤報解決方案

## 問題說明
Windows 防毒軟體可能會將我們的 Odoo-AutoCAD 整合工具誤認為病毒，這是 Python 打包應用程式的常見問題。

## 🔍 為什麼會被誤報？

1. **未簽名的執行檔**：缺少數位簽章
2. **Python 打包特徵**：PyInstaller 打包的共同特徵
3. **新檔案缺乏信譽**：防毒軟體沒有足夠的信譽資料
4. **系統權限需求**：需要存取檔案系統和網路

## ✅ 解決方案

### 1. 立即解決方案

#### Windows Defender
```cmd
# 1. 開啟 Windows 安全性
# 2. 病毒與威脅防護 → 管理設定
# 3. 新增或移除排除項目
# 4. 新增排除項目 → 資料夾
# 5. 選擇: C:\odoo\
```

#### 其他防毒軟體
- **Norton**: 設定 → 防毒 → 掃描和風險 → 排除項目
- **McAfee**: 設定 → 即時掃描 → 排除的檔案
- **Avast**: 設定 → 一般 → 例外狀況

### 2. 建置時的改善

#### A. 使用 PyInstaller 優化參數
```bash
# 建議的 PyInstaller 命令
pyinstaller --onefile \
           --windowed \
           --name "odoo-autocad-integration" \
           --icon "icon/odoo_autocad.ico" \
           --add-data "fonts;fonts" \
           --add-data "icon;icon" \
           --hidden-import "customtkinter" \
           --hidden-import "win32com.client" \
           --exclude-module "pytest" \
           --exclude-module "unittest" \
           --upx-dir "C:/upx" \
           odoo.py
```

#### B. 程式碼簽章（建議）
```bash
# 使用 SignTool (需要程式碼簽章憑證)
signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com output/odoo.exe

signtool sign /f "codesign.pfx" /p "%CODESIGN_PASSWORD%" /t http://timestamp.digicert.com /fd sha256 ../installer/odoo-autocad-integration-3.0-setup.exe

```

### 3. 發佈時的最佳實務

#### A. 建立白名單提交
- 向主要防毒廠商提交檔案進行白名單審查
- **VirusTotal**: https://www.virustotal.com/
- **Microsoft Defender**: https://www.microsoft.com/wdsi/filesubmission

#### B. 提供安裝指南
```markdown
## 安裝前準備
1. 暫時停用即時防護
2. 安裝完成後重新啟用
3. 將安裝資料夾加入排除清單
```

### 4. 程式碼改善措施

#### A. 減少可疑行為
- 避免動態產生執行檔
- 減少系統呼叫
- 使用明確的檔案路徑

#### B. 加入驗證機制
```python
# 檔案完整性檢查
import hashlib

def verify_file_integrity(file_path, expected_hash):
    """驗證檔案完整性"""
    with open(file_path, 'rb') as f:
        file_hash = hashlib.sha256(f.read()).hexdigest()
    return file_hash == expected_hash
```

## 📋 使用者安裝指南

### 步驟 1: 下載前準備
1. 確保從官方來源下載
2. 檢查檔案雜湊值（如提供）

### 步驟 2: 安裝過程
1. 暫時停用即時防護
2. 執行安裝程式
3. 如出現警告，選擇「更多資訊」→「仍要執行」

### 步驟 3: 安裝後設定
1. 重新啟用即時防護
2. 將 `C:\odoo\` 加入排除清單
3. 執行第一次程式測試

## 🏢 企業部署建議

### 1. 群組原則設定
```xml
<!-- Windows Defender 排除設定 -->
<policy>
  <exclusions>
    <folder>C:\odoo\</folder>
    <process>odoo.exe</process>
  </exclusions>
</policy>
```

### 2. 內部憑證簽章
- 使用企業內部 CA 簽發憑證
- 部署企業根憑證到所有電腦

### 3. 段階性部署
1. 先在測試環境驗證
2. 小範圍試點部署
3. 全面推廣

## 📞 技術支援

如果仍遇到誤報問題，請聯繫：
- **Email**: support@your-company.com
- **Phone**: +886-xxx-xxx-xxx

提供以下資訊：
- 防毒軟體名稱和版本
- 錯誤訊息截圖
- Windows 版本
- 安裝日誌檔案