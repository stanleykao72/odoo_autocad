# 除錯模式 Skill

---
name: debug-patterns
description: WPF/C# 應用除錯策略、常見問題診斷與修復模式
trigger-keywords:
  - 除錯
  - debug
  - 例外
  - crash
  - 堆疊追蹤
  - StackTrace
  - Exception
  - 當機
  - 錯誤
  - error
allowed-tools:
  - Read
  - Grep
  - Glob
  - Bash
  - Edit
---

## 概述

本 Skill 提供系統化的除錯策略，用於診斷 WPF 應用、COM Interop 和 API 整合的常見問題。

## 常見問題分類

### 1. WPF 啟動崩潰

**症狀**: 應用程式啟動時立即崩潰

**檢查清單**:
- [ ] XAML 資源載入順序（ResourceDictionary 順序）
- [ ] DI 容器註冊是否完整
- [ ] 建構子是否拋出 `NotImplementedException`
- [ ] 字型資源是否可用

**診斷指令**:
```bash
# 搜尋 NotImplementedException
"C:\Program Files\dotnet\dotnet.exe" build 2>&1 | head -50

# 搜尋未註冊的服務
```

### 2. 導航失敗

**症狀**: 點擊按鈕無反應或白畫面

**檢查清單**:
- [ ] 頁面類別名稱符合 `{Name}Page` 模式
- [ ] Assembly 反射路徑正確
- [ ] DI 中已註冊頁面和 ViewModel
- [ ] NavigationService.Frame 已設定

### 3. COM 連線失敗

**症狀**: COMException 或連線超時

**檢查清單**:
- [ ] AutoCAD 是否已啟動
- [ ] STA 線程合規性
- [ ] COM 物件是否已釋放
- [ ] ProgID 是否正確

### 4. 資料綁定失效

**症狀**: UI 不更新或顯示空白

**檢查清單**:
- [ ] Binding Path 名稱正確（大小寫）
- [ ] DataContext 已設定
- [ ] ViewModel 屬性使用 `[ObservableProperty]`
- [ ] 類別標記為 `partial`

### 5. 建置錯誤

**症狀**: 編譯失敗

**檢查清單**:
- [ ] NuGet 套件已還原
- [ ] 使用完整路徑 `"C:\Program Files\dotnet\dotnet.exe"`
- [ ] 命名空間導入正確
- [ ] 專案參考完整

### 6. 關閉程序掛起

**症狀**: 應用程式無法正常關閉

**檢查清單**:
- [ ] COM 物件已釋放
- [ ] 背景線程已停止
- [ ] CancellationToken 超時設定（5 秒）
- [ ] 每個清理步驟獨立 try-catch

## 除錯優先順序

1. 先看錯誤訊息和堆疊追蹤
2. 確認最近的程式碼變更
3. 驗證建置是否通過
4. 檢查 DI 註冊完整性
5. 搜尋類似問題的已知修復
