# AutoCAD COM Interop Skill

---
name: autocad-com
description: AutoCAD COM 連線、繪圖操作、參數提取的開發指引
trigger-keywords:
  - COM
  - AutoCAD
  - Interop
  - Drawing
  - Marshal
  - GetActiveObject
  - COMException
  - STA
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
  - Bash
---

## 概述

本 Skill 提供 AutoCAD COM Interop 的標準開發模式，確保線程安全與正確的 COM 物件生命週期管理。

## COM 連線模板

```csharp
using System.Runtime.InteropServices;

public class AutoCADService : IAutoCADService, IDisposable
{
    private dynamic? _acadApp;
    private bool _disposed;

    public bool IsConnected => _acadApp != null;

    public bool Connect()
    {
        try
        {
            _acadApp = Marshal.GetActiveObject("AutoCAD.Application");
            return true;
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
        {
            // MK_E_UNAVAILABLE - AutoCAD 未執行
            return false;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (_acadApp != null)
            {
                Marshal.ReleaseComObject(_acadApp);
                _acadApp = null;
            }
            _disposed = true;
        }
    }
}
```

## 線程安全規則

1. COM 操作必須在 **STA 線程**中執行
2. WPF 主線程為 STA，可直接操作
3. 背景線程需使用 `Dispatcher.Invoke` 回傳結果
4. 長時間 COM 操作使用專用 STA 線程：

```csharp
// 專用 STA 線程模式
var result = await Task.Run(() =>
{
    Thread thread = new Thread(() =>
    {
        // COM 操作
    });
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    thread.Join();
});
```

## Python 舊系統對照

| Python (pywin32) | C# (.NET) |
|-------------------|-----------|
| `win32com.client.Dispatch("AutoCAD.Application")` | `Marshal.GetActiveObject("AutoCAD.Application")` |
| 自動垃圾回收 | `Marshal.ReleaseComObject()` 手動釋放 |
| 動態型別 | `dynamic` 關鍵字 |
| `try/except` | `try/catch (COMException)` |

## 常見 COM 錯誤代碼

| HResult | 名稱 | 說明 |
|---------|------|------|
| `0x800401E3` | MK_E_UNAVAILABLE | AutoCAD 未執行 |
| `0x80010001` | RPC_E_CALL_REJECTED | COM 呼叫被拒絕（忙碌） |
| `0x8001010D` | RPC_E_SERVERCALL_RETRYLATER | 伺服器忙碌，稍後重試 |

## 檢查清單

- [ ] COM 物件正確釋放
- [ ] COMException 分類處理
- [ ] STA 線程合規
- [ ] 連線狀態檢查在操作前執行
- [ ] Dispose 模式正確實作
