# CJK 國際化 Skill

---
name: i18n-cjk
description: CJK 字型配置、中日韓文字顯示與本地化支援
trigger-keywords:
  - CJK
  - 字型
  - 中文
  - 本地化
  - 翻譯
  - FontFamily
  - JhengHei
  - 國際化
allowed-tools:
  - Read
  - Write
  - Edit
  - Grep
  - Glob
---

## 概述

本專案需支援正體中文（zh-TW）顯示，使用 CJK 字型鏈確保跨平台相容。

## CJK 字型鏈

```xml
<!-- 標準 CJK 字型鏈（優先順序）-->
FontFamily="Microsoft JhengHei UI, Microsoft YaHei UI, Yu Gothic UI, Segoe UI"
```

| 字型 | 覆蓋語言 | 平台 |
|------|---------|------|
| Microsoft JhengHei UI | 正體中文 | Windows |
| Microsoft YaHei UI | 簡體中文 | Windows |
| Yu Gothic UI | 日文 | Windows |
| Segoe UI | 英文/拉丁字母 (Fallback) | Windows |

## XAML 中的字型設定

### 全域設定
```xml
<!-- App.xaml 或 Resources/Fonts.xaml -->
<FontFamily x:Key="CJKFontFamily">Microsoft JhengHei UI, Microsoft YaHei UI, Yu Gothic UI, Segoe UI</FontFamily>

<Style TargetType="TextBlock">
    <Setter Property="FontFamily" Value="{StaticResource CJKFontFamily}"/>
</Style>
```

### 視窗標題
```xml
<Window FontFamily="Microsoft JhengHei UI, Segoe UI"
        Title="Odoo AutoCAD 整合工具">
```

## 字串本地化（未來擴展）

```csharp
// Resources/Strings.resx (預設 zh-TW)
// Resources/Strings.en-US.resx (英文)

// 使用方式
var title = Properties.Resources.MainWindowTitle;
```

## 檢查清單

- [ ] 所有 UI 文字使用 CJK 字型鏈
- [ ] 視窗標題包含正體中文字型
- [ ] 按鈕、標籤文字正確顯示中文
- [ ] 字型 Fallback 順序正確
