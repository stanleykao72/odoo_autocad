# CLAUDE.md — C# WPF 開發規範

> **層級**: csharp（模組層級）
> **父層級**: 根 CLAUDE.md
> **適用目錄**: `csharp/OdooAutoCADIntegration/`

---

## 專案結構

```
csharp/OdooAutoCADIntegration/
├── OdooAutoCADIntegration.sln          # 解決方案檔案
├── src/
│   ├── OdooAutoCAD.App/                # 主應用程式 (WPF)
│   │   ├── App.xaml / App.xaml.cs      # 應用程式入口
│   │   ├── ViewModels/                 # MVVM ViewModels
│   │   ├── Views/
│   │   │   ├── MainWindow.xaml         # 主視窗
│   │   │   └── Pages/                  # 各功能頁面
│   │   ├── Services/                   # 服務層 (NavigationService 等)
│   │   ├── Converters/                 # XAML 值轉換器
│   │   ├── Resources/                  # 資源字典、圖示
│   │   └── Models/                     # 資料模型
│   ├── OdooAutoCAD.Core/              # 核心業務邏輯（未來）
│   └── OdooAutoCAD.MCP/              # MCP 整合（未來）
└── tests/
    └── OdooAutoCAD.Integration.Tests/ # 整合測試
```

---

## CommunityToolkit.Mvvm 模式

### ViewModel 標準寫法
```csharp
public partial class ExampleViewModel : ObservableObject
{
    private readonly INavigationService _navigationService;

    // 使用 [ObservableProperty] 自動產生屬性
    [ObservableProperty]
    private string _title = string.Empty;

    // 使用 [RelayCommand] 自動產生 ICommand
    [RelayCommand]
    private void DoSomething()
    {
        // 業務邏輯
    }

    // 建構子注入服務
    public ExampleViewModel(INavigationService navigationService)
    {
        _navigationService = navigationService;
    }
}
```

### 命名規則
| 元素 | 命名規則 | 範例 |
|------|---------|------|
| ViewModel 類別 | `{Name}ViewModel` | `MainViewModel` |
| Page 類別 | `{Name}Page` | `DashboardPage` |
| 私有欄位 | `_camelCase` | `_navigationService` |
| ObservableProperty 欄位 | `_camelCase` (自動產生 PascalCase 屬性) | `_title` → `Title` |
| RelayCommand 方法 | `PascalCase` (自動產生 `{Name}Command`) | `DoSomething()` → `DoSomethingCommand` |
| 介面 | `I{Name}` | `INavigationService` |
| 轉換器 | `{Purpose}Converter` | `EqualityConverter` |
| 服務 | `{Name}Service` | `NavigationService` |

---

## WPF / XAML 規範

### 頁面結構
- 每個功能頁面放在 `Views/Pages/{Name}Page.xaml`
- 對應 ViewModel 在 `ViewModels/{Name}ViewModel.cs`
- NavButton Tag 使用 `Btn{PageName}` 模式對應 active indicator

### XAML 撰寫規範
```xml
<!-- 命名空間排列順序 -->
<Page xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:vm="clr-namespace:OdooAutoCAD.App.ViewModels"
      mc:Ignorable="d">

<!-- DataContext 綁定 -->
<Page.DataContext>
    <vm:ExampleViewModel/>
</Page.DataContext>
```

### 資源字典
- 全域樣式放在 `Resources/` 目錄
- 在 `App.xaml` 中合併資源字典
- CJK 字型鏈: `Microsoft JhengHei UI, Microsoft YaHei UI, Yu Gothic UI, Segoe UI`

---

## DI（依賴注入）配置

```csharp
// App.xaml.cs 中使用 Microsoft.Extensions.Hosting
var host = Host.CreateDefaultBuilder()
    .ConfigureServices((context, services) =>
    {
        // 服務註冊
        services.AddSingleton<INavigationService, NavigationService>();

        // ViewModel 註冊
        services.AddTransient<MainViewModel>();
        services.AddTransient<DashboardViewModel>();

        // 視窗註冊
        services.AddSingleton<MainWindow>();
    })
    .Build();
```

### DI 規則
- 服務使用 `AddSingleton` 或 `AddScoped`
- ViewModel 使用 `AddTransient`（每次導航建立新實例）
- 禁止在 ViewModel 中直接 `new` 服務物件

---

## NavigationService 機制

```csharp
// 頁面路由使用反射解析
// typeName 格式: "OdooAutoCAD.App.Views.Pages.{Name}Page"
typeof(NavigationService).Assembly.GetType(typeName);
```

- 導航命令在 `MainViewModel` 中定義
- 頁面實例由 DI 容器管理
- 支援導航歷史（Back/Forward）

---

## 建置指令

```bash
# 完整建置
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln

# 單一專案建置
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/src/OdooAutoCAD.App/OdooAutoCAD.App.csproj

# 執行測試
"C:\Program Files\dotnet\dotnet.exe" test csharp/OdooAutoCADIntegration/tests/OdooAutoCAD.Integration.Tests/ --verbosity normal

# 清理並重建
"C:\Program Files\dotnet\dotnet.exe" clean csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln
"C:\Program Files\dotnet\dotnet.exe" build csharp/OdooAutoCADIntegration/OdooAutoCADIntegration.sln
```

### 建置注意事項
- `dotnet` 不在 PATH 中，必須使用完整路徑
- 已知警告（可忽略）：Core 專案 CS8602 null deref、MCP 專案 CS1998 async
- Integration.Tests 專案需要 `UseWPF=true`
- bash 中使用 `cp` 而非 `copy`

---

## 測試規範

### 測試專案結構
```
tests/OdooAutoCAD.Integration.Tests/
├── ShutdownCleanupTests.cs    # 關閉清理測試
├── NavigationTests.cs         # 導航測試（未來）
├── ViewModelTests.cs          # ViewModel 測試（未來）
└── ServiceTests.cs            # 服務測試（未來）
```

### TDD 工作流
1. **Red**: 寫一個描述期望行為的失敗測試
2. **Green**: 實作最少的程式碼使測試通過
3. **Refactor**: 在測試通過的前提下改善程式碼結構
4. **Build**: 確認整體建置通過
5. **Commit**: 提交（行為變更和結構變更分開提交）

---

## 關閉處理（Shutdown Hardening）

```csharp
// 每個步驟獨立 try-catch，5 秒超時
private async Task CleanupAsync()
{
    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
    try { /* Step 1 */ } catch (Exception ex) { _logger.LogError(ex, "Step 1 failed"); }
    try { /* Step 2 */ } catch (Exception ex) { _logger.LogError(ex, "Step 2 failed"); }
}
```

---

## 已完成的 Sprint

### Sprint 1: App Shell & Navigation (5 batches)
- CJK 字型鏈、視窗圖示
- 7 個佔位頁面
- NavigationService + MainViewModel 導航線路
- Active page indicator (EqualityConverter + DataTrigger)
- Shutdown hardening (per-step try-catch, 5s timeout)
- 6 個 shutdown cleanup 測試
