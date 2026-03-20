# COM 連接策略 — AutoCAD Application 物件取得

> **更新日期**: 2026-03-20
> **影響檔案**: `utility/util_autocad.py` (`connect_autocad` 方法)

## 問題背景

使用 pywin32 連接到已執行的 AutoCAD 時，`GetActiveObject` 或 `client.Dispatch`
取得的 `CDispatch` 物件經常無法存取 `.ActiveDocument`、`.Name` 等屬性，
報錯為 `AttributeError`。

### 根本原因

`GetActiveObject` 透過 ROT (Running Object Table) 取得的是 COM proxy 物件。
當 AutoCAD 忙碌時，COM 呼叫會被 **`RPC_E_CALL_REJECTED`** (HRESULT `-2147418111`)
拒絕。但 pywin32 的 `CDispatch` 封裝會把這個 COM 錯誤包裝成 Python 的
`AttributeError`，導致無法區分「屬性不存在」和「COM 呼叫被拒」。

### 驗證過程

| 策略 | 結果 |
|---|---|
| `client.GetActiveObject()` → `.ActiveDocument` | `AttributeError` (實為 RPC_E_CALL_REJECTED) |
| `client.Dispatch()` → `.ActiveDocument` | 同上 |
| `client.dynamic.Dispatch()` → `.ActiveDocument` | 同上 |
| `client.gencache.EnsureDispatch()` | `TypeError: cannot automate makepy` |
| `pythoncom.CoCreateInstance()` → 低階 `Invoke` | 成功，但建立**新實例**而非連接到現有 AutoCAD |
| `pythoncom.GetActiveObject()` → `QueryInterface` → `Invoke` | 成功，連接到現有實例 |

**結論**：低階 `pythoncom` API 直接呼叫 `GetIDsOfNames` + `Invoke`
可以繞過 `CDispatch` 封裝的問題，正確處理 `RPC_E_CALL_REJECTED` 並重試。

## 目前的連接策略 (v6.0.0.16)

```
pythoncom.GetActiveObject()          ← 取得 PyIUnknown (連接到現有 AutoCAD)
    ↓
raw.QueryInterface(IID_IDispatch)    ← 取得 PyIDispatch
    ↓
idisp.GetIDsOfNames("ActiveDocument") ← 取得 DISPID
    ↓
idisp.Invoke(dispid, PROPERTYGET)    ← 低階呼叫，可正確捕獲 RPC_E_CALL_REJECTED
    ↓
client.Dispatch(idisp)               ← 成功後才包裝成高階 CDispatch
client.Dispatch(doc_idisp)           ← doc 也一起包裝
```

### 重試機制

- 最多 **10 次**重試，每次間隔 **1 秒**
- 捕獲 `RPC_E_CALL_REJECTED` (`-2147418111`) 時自動重試
- 其他 `com_error` 或 `AttributeError` 也會重試

### 程式碼位置

```python
# utility/util_autocad.py — connect_autocad()
raw = pythoncom.GetActiveObject(self.autocad_progid)
for attempt in range(retry_count):
    try:
        idisp = raw.QueryInterface(pythoncom.IID_IDispatch)
        dispid = idisp.GetIDsOfNames(0, "ActiveDocument")
        doc_idisp = idisp.Invoke(dispid, 0, 2, True)  # DISPATCH_PROPERTYGET=2
        if doc_idisp:
            self.acad = client.Dispatch(idisp)
            self.doc = client.Dispatch(doc_idisp)
            break
    except pythoncom.com_error as e:
        time.sleep(retry_delay)
```

## 不可使用的策略

| 策略 | 問題 |
|---|---|
| `CoCreateInstance` | 建立新 AutoCAD 實例，無法取得使用者已開啟的圖檔 |
| `GetActiveObject` + `CDispatch` 直接存取屬性 | `RPC_E_CALL_REJECTED` 被包裝成 `AttributeError` |
| `dir(self.acad)` 觸發型別庫快取 | 不穩定，有時有效有時無效 |
| `gencache.EnsureDispatch` | AutoCAD 型別庫無法自動生成 |

## 其他相關修正 (同次更新)

### Layout 輪詢
- **COM 模式不再輪詢 layout**（`form_main_modern.py` `start_layout_polling()`）
- 僅 IPC 模式每 10 秒偵測 layout 切換
- 原因：COM 模式輪詢會產生大量 `Active Layout Name: Model` 日誌

### MCP Server 日誌
- uvicorn 配置加入 `log_config=None`（`util_mcp_manager.py`）
- 原因：uvicorn 預設 logging config 與應用程式衝突，導致 `Unable to configure formatter 'default'`
