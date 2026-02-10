# WPF 導航系統 Skill

---
name: wpf-navigation
description: NavigationService 頁面路由、導航命令、Active Indicator 機制
trigger-keywords:
  - NavigationService
  - 導航
  - 頁面切換
  - Navigate
  - CurrentPage
  - routing
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
---

## 概述

本專案使用自訂 NavigationService 搭配 Assembly 反射進行頁面路由。

## NavigationService 機制

```csharp
public class NavigationService : INavigationService
{
    private readonly IServiceProvider _serviceProvider;
    private Frame? _frame;

    public void NavigateTo(string pageName)
    {
        // 使用反射解析頁面類型
        var typeName = $"OdooAutoCAD.App.Views.Pages.{pageName}Page";
        var pageType = typeof(NavigationService).Assembly.GetType(typeName);

        if (pageType != null)
        {
            var page = _serviceProvider.GetService(pageType);
            _frame?.Navigate(page);
        }
    }
}
```

## 新增頁面步驟

1. **建立 Page XAML**: `Views/Pages/{Name}Page.xaml`
2. **建立 ViewModel**: `ViewModels/{Name}ViewModel.cs`
3. **註冊 DI**: 在 `App.xaml.cs` 中註冊
4. **加入導航按鈕**: 在 `MainWindow.xaml` 中加入 NavButton
5. **設定 Tag**: NavButton Tag = `Btn{Name}`

## Active Page Indicator

使用 `EqualityConverter` 比對 `CurrentPage` 與 NavButton `Tag`：

```xml
<DataTrigger Value="True">
    <DataTrigger.Binding>
        <MultiBinding Converter="{StaticResource EqualityConverter}">
            <Binding Path="CurrentPage"/>
            <Binding Path="Tag" RelativeSource="{RelativeSource Self}"/>
        </MultiBinding>
    </DataTrigger.Binding>
    <Setter Property="Background" Value="{StaticResource ActivePageBrush}"/>
</DataTrigger>
```

## 已有頁面清單

| 頁面 | 檔案 | NavButton Tag |
|------|------|--------------|
| Dashboard | `DashboardPage.xaml` | `BtnDashboard` |
| AutoCAD Parameters | `AutoCADParametersPage.xaml` | `BtnAutoCADParameters` |
| BOQ Management | `BOQManagementPage.xaml` | `BtnBOQManagement` |
| Purchase Requisition | `PurchaseRequisitionPage.xaml` | `BtnPurchaseRequisition` |
| Odoo Sync | `OdooSyncPage.xaml` | `BtnOdooSync` |
| Settings | `SettingsPage.xaml` | `BtnSettings` |
| About | `AboutPage.xaml` | `BtnAbout` |

## 檢查清單

- [ ] 頁面類別名稱符合 `{Name}Page` 模式
- [ ] ViewModel 已在 DI 中註冊
- [ ] NavButton Tag 符合 `Btn{Name}` 模式
- [ ] NavigationService 可反射解析新頁面
