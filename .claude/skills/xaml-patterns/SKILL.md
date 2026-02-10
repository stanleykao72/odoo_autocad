# XAML 模式 Skill

---
name: xaml-patterns
description: WPF XAML 模式、DataTemplate、Style、Converter、Binding 開發指引
trigger-keywords:
  - XAML
  - DataTemplate
  - Style
  - Converter
  - Binding
  - ResourceDictionary
  - DataTrigger
  - ControlTemplate
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
---

## 概述

本 Skill 提供 WPF XAML 的標準模式，涵蓋頁面結構、資料綁定、樣式、轉換器等。

## 頁面結構模板

```xml
<Page x:Class="OdooAutoCAD.App.Views.Pages.{Name}Page"
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      mc:Ignorable="d"
      d:DesignHeight="600" d:DesignWidth="800"
      Title="{Name}Page">

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- 標題 -->
        <TextBlock Grid.Row="0"
                   Text="頁面標題"
                   FontSize="24"
                   FontWeight="Bold"
                   FontFamily="Microsoft JhengHei UI, Segoe UI"
                   Margin="0,0,0,16"/>

        <!-- 內容區域 -->
        <Grid Grid.Row="1">
            <!-- 頁面內容 -->
        </Grid>
    </Grid>
</Page>
```

## 值轉換器模式

```csharp
// Converters/{Purpose}Converter.cs
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is true ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is Visibility.Visible;
    }
}
```

```xml
<!-- XAML 中使用 -->
<Page.Resources>
    <converters:BoolToVisibilityConverter x:Key="BoolToVisibility"/>
</Page.Resources>

<ProgressBar Visibility="{Binding IsLoading, Converter={StaticResource BoolToVisibility}}"/>
```

## 樣式與資源

```xml
<!-- Resources/Styles.xaml -->
<ResourceDictionary>
    <!-- CJK 字型鏈 -->
    <FontFamily x:Key="CJKFontFamily">Microsoft JhengHei UI, Microsoft YaHei UI, Yu Gothic UI, Segoe UI</FontFamily>

    <!-- 按鈕樣式 -->
    <Style x:Key="PrimaryButton" TargetType="Button">
        <Setter Property="Background" Value="#0078D4"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="Padding" Value="16,8"/>
        <Setter Property="FontFamily" Value="{StaticResource CJKFontFamily}"/>
    </Style>
</ResourceDictionary>
```

## 導航按鈕 Active Indicator

```xml
<!-- 使用 EqualityConverter + DataTrigger -->
<RadioButton Tag="BtnDashboard"
             Command="{Binding NavigateCommand}"
             CommandParameter="Dashboard">
    <RadioButton.Style>
        <Style TargetType="RadioButton">
            <Style.Triggers>
                <DataTrigger Value="True">
                    <DataTrigger.Binding>
                        <MultiBinding Converter="{StaticResource EqualityConverter}">
                            <Binding Path="CurrentPage"/>
                            <Binding Path="Tag" RelativeSource="{RelativeSource Self}"/>
                        </MultiBinding>
                    </DataTrigger.Binding>
                    <Setter Property="Background" Value="#E3F2FD"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </RadioButton.Style>
</RadioButton>
```

## 檢查清單

- [ ] 命名空間排列順序正確（WPF → blend → 自訂）
- [ ] CJK 字型鏈已設定
- [ ] Converter 在 Resources 中正確註冊
- [ ] Binding Path 對應 ViewModel 屬性名稱
- [ ] 使用 `{StaticResource}` 而非硬編碼值
