# Odoo 整合模式規範

> **優先級**: MEDIUM
> **適用**: Odoo API 開發者、AutoCAD 整合專家

---

## 1. API 客戶端模式

### HttpClient 管理
```csharp
// ✅ 使用 IHttpClientFactory
services.AddHttpClient<IOdooApiClient, OdooApiClient>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
});

// ❌ 直接建立 HttpClient（socket exhaustion）
var client = new HttpClient();
```

### JSON-RPC 呼叫格式
```csharp
// Odoo JSON-RPC 標準格式
var request = new
{
    jsonrpc = "2.0",
    method = "call",
    @params = new { /* 參數 */ }
};
```

## 2. 認證流程

1. POST `/web/session/authenticate` 取得 session
2. 後續請求帶上 session cookie
3. Session 過期時自動重新認證
4. 不快取密碼（只快取 session）

## 3. 資料同步策略

| 操作 | 方法 | 說明 |
|------|------|------|
| 查詢 | `search_read` | 讀取 Odoo 資料 |
| 建立 | `create` | 新增記錄 |
| 更新 | `write` | 修改記錄 |
| 刪除 | `unlink` | 刪除記錄 |

### 同步原則
- 讀取操作可快取（TTL 依資料類型）
- 寫入操作必須即時
- 批次操作優先（減少 API 呼叫次數）
- 失敗操作需要重試機制

## 4. Python 舊系統對照

遷移時參考 Python 版本的：
- `utility/util_odoo.py` — API 呼叫模式
- `utility/util_push_to_boq.py` — BOQ 處理邏輯
- `config/*.yaml` — 配置格式

**Python → C# 對應表**:
| Python | C# |
|--------|-----|
| `bravado` client | `HttpClient` + JSON-RPC |
| `yaml.load()` | `IConfiguration` |
| `SQLAlchemy` | `Entity Framework Core` |
| `dict` response | `JsonDocument` / 強型別模型 |

## 5. 錯誤處理層級

```
API 呼叫
├── HttpRequestException → 網路層錯誤
│   ├── 401 → 重新認證
│   ├── 5xx → 重試（最多 3 次）
│   └── 其他 → 記錄並拋出
├── TaskCanceledException → 超時
│   └── 記錄並通知使用者
└── JsonException → 回應解析錯誤
    └── 記錄原始回應並拋出
```

## 檢查清單

- [ ] 使用 IHttpClientFactory
- [ ] 認證資訊不硬編碼
- [ ] Session 過期自動重新認證
- [ ] 錯誤分層處理
- [ ] 寫入操作即時執行
- [ ] 批次操作優先
