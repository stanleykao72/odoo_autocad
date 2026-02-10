# TASK-FE-{xxx}-{yy}: {前端任務標題}

> **狀態**: PENDING | IN_PROGRESS | REVIEW | TESTING | DONE | BLOCKED
> **指派代理**: WPF 開發者
> **來源用戶故事**: US-{xxx}-{yy}
> **建立日期**: {YYYY-MM-DD}

---

## 任務描述

{詳細描述此 UI/前端任務}

## UI 規格

### 頁面/元件

| 元件 | 檔案 | 說明 |
|------|------|------|
| `{Name}Page` | `Views/Pages/{Name}Page.xaml` | {頁面說明} |
| `{Name}ViewModel` | `ViewModels/{Name}ViewModel.cs` | {ViewModel 說明} |

### 版面配置

```
┌──────────────────────────────────┐
│ 標題區域                          │
├──────────────────────────────────┤
│                                  │
│ 內容區域                          │
│                                  │
├──────────────────────────────────┤
│ 動作按鈕區域                      │
└──────────────────────────────────┘
```

### 資料綁定

| UI 元素 | Binding Path | 模式 |
|---------|-------------|------|
| {元素} | `{Property}` | OneWay/TwoWay |

### 命令

| 命令 | 觸發方式 | 行為 |
|------|---------|------|
| `{Name}Command` | Button Click | {行為描述} |

## 樣式需求

- [ ] CJK 字型鏈已設定
- [ ] 使用 StaticResource（非硬編碼值）
- [ ] 響應式佈局（Grid RowDefinitions）
- [ ] 一致的 Margin/Padding

## 導航整合

- NavButton Tag: `Btn{Name}`
- NavigationService 路由: `{Name}Page`
- DI 註冊: `services.AddTransient<{Name}Page>()`

## 依賴

| 依賴項 | 狀態 |
|--------|------|
| {ViewModel 服務依賴} | 已完成/待處理 |
| {API 服務} | 已完成/待處理 |

## 交付標準

- [ ] XAML 頁面完成
- [ ] ViewModel 完成並綁定
- [ ] 導航可正常切換
- [ ] CJK 字型顯示正確
- [ ] 建置無新增錯誤
- [ ] DI 註冊完成
- [ ] 準備交接給 QA
