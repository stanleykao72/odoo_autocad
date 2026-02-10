# Odoo REST API Skill

---
name: odoo-rest-api
description: Odoo REST API 認證、資料查詢、同步的 C# 整合指引
trigger-keywords:
  - Odoo API
  - REST
  - 同步
  - 認證
  - authenticate
  - search_read
  - JsonRpc
  - HttpClient
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
  - Bash
---

## 概述

本 Skill 提供 Odoo REST API 的 C# 整合模式，涵蓋認證、CRUD 操作與資料同步。

## Odoo JSON-RPC 協議

### 認證
```csharp
public async Task<bool> AuthenticateAsync(string url, string db, string login, string password)
{
    var request = new
    {
        jsonrpc = "2.0",
        method = "call",
        @params = new { db, login, password }
    };

    var response = await _httpClient.PostAsJsonAsync($"{url}/web/session/authenticate", request);
    var result = await response.Content.ReadFromJsonAsync<JsonDocument>();

    // 檢查 session_id cookie
    return response.IsSuccessStatusCode && result?.RootElement.GetProperty("result").GetProperty("uid").GetInt32() > 0;
}
```

### 資料查詢 (search_read)
```csharp
public async Task<List<T>> SearchReadAsync<T>(string model, object[] domain, string[] fields, int limit = 0)
{
    var request = new
    {
        jsonrpc = "2.0",
        method = "call",
        @params = new
        {
            model,
            method = "search_read",
            args = new object[] { domain },
            kwargs = new { fields, limit }
        }
    };

    var response = await _httpClient.PostAsJsonAsync($"{_url}/web/dataset/call_kw", request);
    // 解析回應...
}
```

## Python 舊系統對照

| Python (Bravado) | C# (HttpClient) |
|-------------------|-----------------|
| `swagger_client.authenticate()` | `HttpClient.PostAsJsonAsync` |
| `client.search_read()` | 自訂 `SearchReadAsync` |
| `yaml.load()` 配置 | `IConfiguration` + JSON |
| SQLAlchemy ORM | Entity Framework Core |

## 資料同步策略

| 資料類型 | 策略 | 頻率 |
|---------|------|------|
| 產品清單 | 完整同步 + 差異更新 | 啟動時 + 按需 |
| 專案資料 | 按需查詢 | 使用時 |
| BOQ | 即時推送 | 變更時 |
| PR | 批次提交 | 確認時 |

## 錯誤處理模式

```csharp
// 分層錯誤處理
catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.Unauthorized)
{
    // Session 過期 → 重新認證
}
catch (HttpRequestException ex) when (ex.StatusCode >= HttpStatusCode.InternalServerError)
{
    // 伺服器錯誤 → 記錄並重試
}
catch (TaskCanceledException)
{
    // 請求超時
}
```

## 檢查清單

- [ ] 使用 `IHttpClientFactory` 管理 HttpClient
- [ ] 認證資訊不硬編碼
- [ ] API 回應正確驗證
- [ ] 網路錯誤有重試策略
- [ ] 超時設定合理（30s 預設）
