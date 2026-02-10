# C# MVVM 模式 Skill

---
name: csharp-mvvm
description: CommunityToolkit.Mvvm 模式與 ViewModel 開發指引
trigger-keywords:
  - ObservableProperty
  - RelayCommand
  - ViewModel
  - MVVM
  - ObservableObject
  - NotifyPropertyChanged
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
---

## 概述

本 Skill 提供 CommunityToolkit.Mvvm 的標準開發模式，確保所有 ViewModel 遵循一致的架構。

## 核心模式

### ViewModel 基本結構

```csharp
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OdooAutoCAD.App.ViewModels;

public partial class {Name}ViewModel : ObservableObject
{
    // 1. 私有服務欄位
    private readonly INavigationService _navigationService;

    // 2. Observable 屬性（使用 source generator）
    [ObservableProperty]
    private string _title = string.Empty;

    [ObservableProperty]
    private bool _isLoading;

    // 3. 建構子（DI 注入）
    public {Name}ViewModel(INavigationService navigationService)
    {
        _navigationService = navigationService;
    }

    // 4. 命令方法
    [RelayCommand]
    private void Navigate(string page)
    {
        _navigationService.NavigateTo(page);
    }

    [RelayCommand]
    private async Task LoadDataAsync()
    {
        IsLoading = true;
        try
        {
            // 載入資料
        }
        finally
        {
            IsLoading = false;
        }
    }
}
```

### 屬性連動通知

```csharp
[ObservableProperty]
[NotifyPropertyChangedFor(nameof(FullName))]
private string _firstName = string.Empty;

[ObservableProperty]
[NotifyPropertyChangedFor(nameof(FullName))]
private string _lastName = string.Empty;

public string FullName => $"{FirstName} {LastName}";
```

### 命令可執行條件

```csharp
[RelayCommand(CanExecute = nameof(CanSave))]
private void Save()
{
    // 儲存邏輯
}

private bool CanSave => !string.IsNullOrEmpty(Title);
```

## 檢查清單

- [ ] 類別標記為 `partial`
- [ ] 繼承 `ObservableObject`
- [ ] 私有欄位使用 `_camelCase`（source generator 產生 `PascalCase` 屬性）
- [ ] 服務透過建構子注入
- [ ] 非同步命令使用 `async Task`（非 `async void`）
- [ ] 命令方法名稱不含 `Command` 後綴（自動產生）

## 常見錯誤

| 錯誤 | 修正 |
|------|------|
| 類別未標記 `partial` | 加上 `partial` 關鍵字 |
| 手動實作 INotifyPropertyChanged | 使用 `[ObservableProperty]` |
| 在 ViewModel 中 `new` 服務 | 透過建構子 DI 注入 |
| `async void` 命令方法 | 改為 `async Task` |
| 在命令中忽略例外 | 加上 try-catch 並通知 UI |
