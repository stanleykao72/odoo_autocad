# DI 依賴注入模式 Skill

---
name: di-patterns
description: Microsoft.Extensions.Hosting DI 容器配置與服務注入模式
trigger-keywords:
  - DI
  - 依賴注入
  - IServiceProvider
  - Hosting
  - AddSingleton
  - AddTransient
  - ConfigureServices
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
---

## 概述

本專案使用 `Microsoft.Extensions.Hosting` 進行依賴注入管理。

## DI 容器配置

### App.xaml.cs 中的服務註冊

```csharp
var host = Host.CreateDefaultBuilder()
    .ConfigureServices((context, services) =>
    {
        // === 服務層 ===
        services.AddSingleton<INavigationService, NavigationService>();
        // services.AddSingleton<IAutoCADService, AutoCADService>();  // 未來
        // services.AddSingleton<IOdooApiClient, OdooApiClient>();    // 未來

        // === ViewModel 層 ===
        services.AddTransient<MainViewModel>();
        services.AddTransient<DashboardViewModel>();
        // 其他 ViewModel...

        // === 視窗/頁面 ===
        services.AddSingleton<MainWindow>();
        services.AddTransient<DashboardPage>();
        // 其他頁面...
    })
    .Build();
```

## 服務生命週期規則

| 生命週期 | 適用場景 | 範例 |
|---------|---------|------|
| `Singleton` | 全域唯一服務 | NavigationService, AutoCADService |
| `Transient` | 每次建立新實例 | ViewModel, Page |
| `Scoped` | 請求範圍內唯一 | 不常用於 WPF |

## 注入模式

### 建構子注入（推薦）
```csharp
// ✅ 正確
public class MainViewModel : ObservableObject
{
    private readonly INavigationService _nav;
    public MainViewModel(INavigationService nav) => _nav = nav;
}
```

### 禁止模式
```csharp
// ❌ 禁止：Service Locator
var service = App.ServiceProvider.GetService<INavigationService>();

// ❌ 禁止：直接 new
var nav = new NavigationService();
```

## 新增服務步驟

1. 定義介面 `I{Name}Service`
2. 實作類別 `{Name}Service`
3. 在 `App.xaml.cs` 的 `ConfigureServices` 中註冊
4. 在需要的 ViewModel/Service 建構子中注入

## 檢查清單

- [ ] 服務有對應的介面
- [ ] 生命週期選擇正確
- [ ] 建構子注入（非 Service Locator）
- [ ] 循環依賴已避免
