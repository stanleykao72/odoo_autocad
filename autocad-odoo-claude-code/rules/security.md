# 安全規範

> **優先級**: CRITICAL
> **適用**: 所有代理

---

## 1. 秘密管理

### 規則
所有敏感資訊（密碼、API Key、連線字串、Token）禁止出現在原始碼中。

### 正確做法
```csharp
// ✅ 使用配置檔案
var password = _configuration["Odoo:Password"];

// ✅ 使用環境變數
var apiKey = Environment.GetEnvironmentVariable("ODOO_API_KEY");

// ✅ 使用 .NET User Secrets（開發環境）
// dotnet user-secrets set "Odoo:Password" "secret"
```

### 錯誤做法
```csharp
// ❌ 硬編碼密碼
var password = "admin123";

// ❌ 硬編碼連線字串
var connectionString = "Server=192.168.1.100;Database=odoo;User=admin;Password=secret";

// ❌ 硬編碼 API URL 含 Token
var url = "https://odoo.example.com/api?token=abc123";
```

### 檢查指令
```
Grep: pattern="password\s*=\s*\"[^\"]+\"" path="csharp/" type="cs"
Grep: pattern="api.?key\s*=\s*\"" path="csharp/" type="cs" -i
```

---

## 2. .gitignore 完整性

### 必須包含的排除項目
```
# 秘密檔案
appsettings.Development.json
appsettings.Local.json
*.pfx
*.key

# 使用者密鑰
secrets.json

# 環境變數
.env
.env.local
```

---

## 3. COM 物件安全

### 規則
所有 COM 物件必須正確釋放，避免記憶體洩漏和行程殘留。

### 正確做法
```csharp
// ✅ 使用 try-finally 確保釋放
dynamic? comObj = null;
try
{
    comObj = Marshal.GetActiveObject("AutoCAD.Application");
    // 操作...
}
finally
{
    if (comObj != null)
        Marshal.ReleaseComObject(comObj);
}

// ✅ 實作 IDisposable
public class AutoCADService : IDisposable
{
    public void Dispose() { /* 釋放 COM 物件 */ }
}
```

### 錯誤做法
```csharp
// ❌ 未釋放 COM 物件
var acad = Marshal.GetActiveObject("AutoCAD.Application");
// 使用後忘記釋放

// ❌ 在 finalizer 中釋放（不可靠）
~AutoCADService() { Marshal.ReleaseComObject(_acadApp); }
```

---

## 4. HTTP 通訊安全

### 規則
- 所有 API 通訊必須使用 HTTPS
- 驗證 SSL 憑證
- 設定請求超時
- 不在日誌中記錄敏感資訊

### 正確做法
```csharp
// ✅ HTTPS + 超時
var client = new HttpClient
{
    BaseAddress = new Uri("https://odoo.example.com"),
    Timeout = TimeSpan.FromSeconds(30)
};

// ✅ 日誌不含密碼
_logger.LogInformation("Connecting to Odoo at {Url}", url);
```

### 錯誤做法
```csharp
// ❌ 使用 HTTP
var client = new HttpClient { BaseAddress = new Uri("http://odoo.example.com") };

// ❌ 停用 SSL 驗證
handler.ServerCertificateCustomValidationCallback = (_, _, _, _) => true;

// ❌ 日誌含密碼
_logger.LogInformation("Login with password: {Password}", password);
```

---

## 5. 輸入驗證

### 規則
所有外部輸入必須驗證：API 回應、檔案路徑、使用者輸入。

```csharp
// ✅ 驗證 API 回應
if (response.StatusCode != HttpStatusCode.OK)
    throw new OdooApiException($"Unexpected status: {response.StatusCode}");

// ✅ 驗證檔案路徑
if (!Path.GetFullPath(filePath).StartsWith(allowedDirectory))
    throw new SecurityException("Path traversal detected");
```

---

## 安全檢查清單

| 項目 | 嚴重性 | 說明 |
|------|--------|------|
| 硬編碼秘密 | CRITICAL | 程式碼中無密碼/Key/Token |
| COM 物件洩漏 | HIGH | 所有 COM 物件正確釋放 |
| HTTPS 強制 | HIGH | API 通訊使用 HTTPS |
| SSL 驗證 | HIGH | 不停用憑證驗證 |
| 日誌安全 | MEDIUM | 日誌不含敏感資訊 |
| 輸入驗證 | HIGH | 外部輸入均已驗證 |
| .gitignore | MEDIUM | 敏感檔案已排除 |
| 錯誤訊息 | MEDIUM | 不向使用者暴露內部細節 |
