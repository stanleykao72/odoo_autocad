# 編碼風格規範

> **優先級**: HIGH
> **適用**: 所有代理

---

## 1. C# 命名規範

| 元素 | 規則 | 範例 |
|------|------|------|
| 命名空間 | PascalCase | `OdooAutoCAD.App.ViewModels` |
| 類別/介面 | PascalCase | `MainViewModel`, `INavigationService` |
| 公開方法 | PascalCase | `NavigateTo()`, `LoadDataAsync()` |
| 私有方法 | PascalCase | `ValidateInput()` |
| 公開屬性 | PascalCase | `Title`, `IsLoading` |
| 私有欄位 | _camelCase | `_navigationService`, `_title` |
| 參數 | camelCase | `pageName`, `isEnabled` |
| 常數 | PascalCase | `MaxRetryCount`, `DefaultTimeout` |
| 非同步方法 | PascalCase + Async | `LoadDataAsync()`, `ConnectAsync()` |

## 2. 檔案組織

### 目錄結構
```
src/OdooAutoCAD.App/
├── ViewModels/          # 所有 ViewModel
├── Views/
│   ├── MainWindow.xaml  # 主視窗
│   └── Pages/           # 所有功能頁面
├── Services/            # 服務層
├── Converters/          # XAML 值轉換器
├── Resources/           # 資源字典、圖示
└── Models/              # 資料模型
```

### using 排列順序
```csharp
// 1. System 命名空間
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

// 2. Microsoft/第三方
using Microsoft.Extensions.DependencyInjection;
using CommunityToolkit.Mvvm.ComponentModel;

// 3. 專案內部
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.Models;
```

## 3. 類別結構順序

```csharp
public partial class ExampleViewModel : ObservableObject
{
    // 1. 常數
    private const int MaxRetries = 3;

    // 2. 私有服務欄位
    private readonly INavigationService _navigationService;

    // 3. Observable 屬性
    [ObservableProperty]
    private string _title = string.Empty;

    // 4. 計算屬性
    public string DisplayTitle => $"[{Title}]";

    // 5. 建構子
    public ExampleViewModel(INavigationService nav) { _navigationService = nav; }

    // 6. 公開方法/命令
    [RelayCommand]
    private void Navigate(string page) { }

    // 7. 私有方法
    private void ValidateInput() { }
}
```

## 4. 程式碼品質規則

### 必須
- 每個類別一個檔案
- 方法長度不超過 30 行（建議）
- 使用 `var` 當型別明顯時
- 非同步方法名稱以 `Async` 結尾
- `IDisposable` 物件使用 `using` 語句

### 禁止
- 空的 catch 區塊（至少記錄日誌）
- `async void`（除了事件處理程序）
- 巢狀深度超過 3 層
- 魔術數字（使用常數）
- 未使用的 `using` 導入

## 5. XAML 格式

```xml
<!-- 屬性排列：每行一個屬性 -->
<Button Content="確定"
        Command="{Binding SaveCommand}"
        Style="{StaticResource PrimaryButton}"
        Margin="0,8,0,0"
        HorizontalAlignment="Right"/>

<!-- 短屬性可寫同一行 -->
<TextBlock Text="標題" FontWeight="Bold"/>
```

## 程式碼品質檢查清單

- [ ] 命名遵循規範
- [ ] 檔案組織正確
- [ ] using 排列正確
- [ ] 類別結構順序正確
- [ ] 無空 catch 區塊
- [ ] 無 async void
- [ ] 無魔術數字
- [ ] XAML 格式整齊
