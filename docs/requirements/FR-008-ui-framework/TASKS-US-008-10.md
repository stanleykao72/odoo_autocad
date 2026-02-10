# TASKS: US-008-10 — CJK Font Rendering

> **Parent US**: [US-008-10](US-008-10-cjk-font-rendering.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 4S + 1M + 0L
> **Status**: Done

## Prerequisites
- [x] None (font configuration is independent of other US stories)

## Acceptance Criteria
- [x] AC-01: The application uses Microsoft JhengHei UI as the primary font family for all UI elements.
- [x] AC-02: A font fallback chain is defined: Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS.
- [x] AC-03: The font family is set at the Window or Application level so it applies globally without per-element configuration.
- [x] AC-04: All Chinese labels, status messages, log entries, and user-facing text render correctly without missing glyphs or tofu characters.
- [x] AC-05: The font family resource is defined as a named resource (`AppFontFamily`) in the resource dictionary for consistent reference.
- [x] AC-06: All styles and templates reference the `AppFontFamily` resource rather than hardcoding font names.

---

## TASK-008-10-01: Define AppFontFamily resource with fallback chain in ResourceDictionary

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-008-10-02, TASK-008-10-04 |

### What to do
- Add a `FontFamily` resource to `App.xaml` ResourceDictionary:
  ```xml
  <FontFamily x:Key="AppFontFamily">Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS</FontFamily>
  ```
- The comma-separated list defines the fallback chain: WPF automatically tries each font in order
- Place this resource near the top of the ResourceDictionary, before styles that reference it
- Optionally define font size resources for consistency:
  ```xml
  <sys:Double x:Key="FontSizeTitle">16</sys:Double>
  <sys:Double x:Key="FontSizeHeading">14</sys:Double>
  <sys:Double x:Key="FontSizeBody">12</sys:Double>
  <sys:Double x:Key="FontSizeSmall">10</sys:Double>
  ```
- Add `xmlns:sys="clr-namespace:System;assembly=mscorlib"` namespace if using `sys:Double`

### How to verify
- [x] AppFontFamily resource is defined with the correct fallback chain (AC-02, AC-05)
- [x] Resource is in the application-level ResourceDictionary (AC-05)

---

## TASK-008-10-02: Set FontFamily at the Window level referencing AppFontFamily

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | TASK-008-10-01 |
| Blocks | None |

### What to do
- Add `FontFamily="{StaticResource AppFontFamily}"` attribute to the `<Window>` element in MainWindow.xaml
- This applies the CJK font family globally to all controls within the window without per-element configuration
- Optionally set `FontSize="12"` at the Window level for consistent base font size
- WPF property value inheritance propagates the FontFamily to all descendant elements unless explicitly overridden

### How to verify
- [x] MainWindow.xaml has FontFamily set to AppFontFamily resource (AC-01, AC-03)
- [x] CJK characters render correctly throughout the window (AC-04)
- [x] No per-element font configuration is needed for basic CJK rendering (AC-03)

---

## TASK-008-10-03: Merge font resources into App.xaml ResourceDictionary

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml` |
| Estimate | S |
| Depends On | TASK-008-10-01 |
| Blocks | None |

### What to do
- If font resources are defined in a separate `Themes/Fonts.xaml` ResourceDictionary, merge it into `App.xaml`:
  ```xml
  <Application.Resources>
      <ResourceDictionary>
          <ResourceDictionary.MergedDictionaries>
              <ResourceDictionary Source="Themes/Fonts.xaml"/>
          </ResourceDictionary.MergedDictionaries>
          <!-- other resources -->
      </ResourceDictionary>
  </Application.Resources>
  ```
- If font resources are defined inline in `App.xaml` (as done in TASK-008-10-01), this merge step is not needed
- Ensure the `AppFontFamily` resource is available to all windows and pages via the application-level dictionary
- Consider splitting resources into separate files (`Themes/Colors.xaml`, `Themes/Styles.xaml`, `Themes/Fonts.xaml`) if the number grows, per FR-008 implementation notes

### How to verify
- [x] Font resources are accessible from any window or page in the application (AC-03)
- [x] `{StaticResource AppFontFamily}` resolves correctly in all XAML files (AC-05)

---

## TASK-008-10-04: Verify all named styles reference AppFontFamily

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml` |
| Estimate | S |
| Depends On | TASK-008-10-01 |
| Blocks | None |

### What to do
- Review `PrimaryButton` style: add `<Setter Property="FontFamily" Value="{StaticResource AppFontFamily}"/>` if not inherited
- Review `NavButton` style: add FontFamily setter if the ControlTemplate overrides inheritance
- Review `CardPanel` style: Border does not have FontFamily, but child TextBlocks inherit from Window
- Review `ModernTextBox` style: add `<Setter Property="FontFamily" Value="{StaticResource AppFontFamily}"/>` since ControlTemplate may break inheritance
- For any style with a `ControlTemplate`, explicitly set FontFamily on the template's content elements or via TemplateBinding to ensure CJK rendering
- Check that no styles hardcode a non-CJK font family (e.g., "Segoe UI" without fallback)

### How to verify
- [x] All named styles reference AppFontFamily resource (AC-06)
- [x] No hardcoded font names exist in styles or templates (AC-06)
- [x] CJK text renders correctly in buttons, text boxes, and card panels (AC-04)

---

## TASK-008-10-05: Test CJK rendering with sample Chinese strings across all UI areas

| Field | Value |
|-------|-------|
| Target | `tests/` (UI test or manual verification) |
| Estimate | M |
| Depends On | TASK-008-10-02, TASK-008-10-04 |
| Blocks | None |

### What to do
- Create a test or verification checklist that displays Chinese text in all UI areas:
  - Sidebar title area: "Odoo AutoCAD" (ASCII) and "整合系統 v6.0" (CJK variant)
  - Navigation buttons: mix of emoji + English (default) or CJK labels
  - Connection status text: "已連接" / "未連接" / "已停止"
  - Header page title: "儀表板" / "AutoCAD 整合" (CJK page titles if localized)
  - Status bar message: "就緒" / "正在重新整理..."
  - System log panel: entries with "[SSE GUI]", "[Odoo]" category prefixes and CJK content
- Verify no tofu characters (empty rectangles) or missing glyphs appear
- Test on a system with Microsoft JhengHei UI installed, and on a system without it to verify fallback chain
- Verify font sizes are consistent between CJK and Latin characters

### How to verify
- [x] Chinese labels render correctly in sidebar, header, status bar, and log panel (AC-04)
- [x] No missing glyphs or tofu characters (AC-04)
- [x] Font fallback chain activates when primary font is unavailable (AC-02)

---

## Dependency Graph

```
TASK-008-10-01 (AppFontFamily resource)
    ├── TASK-008-10-02 (Window-level FontFamily)
    │   └── TASK-008-10-05 (CJK rendering test)
    ├── TASK-008-10-03 (Merge into App.xaml)
    └── TASK-008-10-04 (Verify styles)
        └── TASK-008-10-05 (CJK rendering test)
```
