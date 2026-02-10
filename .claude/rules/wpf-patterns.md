# WPF 開發模式規範

> **優先級**: MEDIUM
> **適用**: WPF 開發者代理、QA 測試工程師

---

## 1. MVVM 嚴格分離

### 規則
- **View（XAML）**: 只負責 UI 呈現，不含業務邏輯
- **ViewModel**: 業務邏輯，透過 Binding 與 View 互動
- **Model**: 資料結構，不依賴 UI

### 正確做法
```csharp
// ✅ ViewModel 處理邏輯
public partial class DashboardViewModel : ObservableObject
{
    [RelayCommand]
    private async Task RefreshAsync()
    {
        var data = await _dataService.LoadAsync();
        Items = new ObservableCollection<ItemModel>(data);
    }
}
```

### 錯誤做法
```csharp
// ❌ Code-behind 中寫業務邏輯
public partial class DashboardPage : Page
{
    private async void Button_Click(object sender, RoutedEventArgs e)
    {
        var data = await LoadFromApi(); // 業務邏輯不應在這裡
        DataGrid.ItemsSource = data;
    }
}
```

## 2. 資料綁定規則

- 使用 `{Binding}` 而非 code-behind 存取
- 集合使用 `ObservableCollection<T>`
- 命令使用 `[RelayCommand]`
- 雙向綁定使用 `{Binding Mode=TwoWay}`

## 3. 資源管理

```xml
<!-- ✅ 使用 StaticResource -->
<TextBlock Style="{StaticResource HeaderStyle}"/>

<!-- ❌ 硬編碼值 -->
<TextBlock FontSize="24" FontWeight="Bold" Foreground="#333"/>
```

### 資源字典合併順序（App.xaml）
```xml
<Application.Resources>
    <ResourceDictionary>
        <ResourceDictionary.MergedDictionaries>
            <!-- 1. 字型 -->
            <ResourceDictionary Source="Resources/Fonts.xaml"/>
            <!-- 2. 色彩 -->
            <ResourceDictionary Source="Resources/Colors.xaml"/>
            <!-- 3. 樣式 -->
            <ResourceDictionary Source="Resources/Styles.xaml"/>
            <!-- 4. 轉換器 -->
            <ResourceDictionary Source="Resources/Converters.xaml"/>
        </ResourceDictionary.MergedDictionaries>
    </ResourceDictionary>
</Application.Resources>
```

## 4. 導航模式

- 頁面位於 `Views/Pages/{Name}Page.xaml`
- NavButton Tag = `Btn{PageName}`
- NavigationService 使用 Assembly 反射
- DI 管理頁面實例

## 5. CJK 字型

**必須**在以下位置設定 CJK 字型鏈：
- 全域樣式（App.xaml 或 ResourceDictionary）
- 視窗標題
- 所有包含中文的 TextBlock/Button

## 6. 關閉處理

- 每個清理步驟獨立 `try-catch`
- 使用 `CancellationToken` + 5 秒超時
- COM 物件在關閉時釋放
- 背景線程在關閉時停止

## 檢查清單

- [ ] 業務邏輯在 ViewModel（非 code-behind）
- [ ] 使用 Binding 而非直接存取 UI 元素
- [ ] 資源使用 StaticResource
- [ ] ResourceDictionary 合併順序正確
- [ ] CJK 字型鏈已設定
- [ ] 關閉處理健壯
