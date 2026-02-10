# TASKS: US-008-05 — Window Default Size

> **Parent US**: [US-008-05](US-008-05-window-default-size.md)
> **Parent FR**: [FR-008](FR-008-ui-framework.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 5S + 0M + 0L
> **Status**: Done

## Prerequisites
- [x] None (this is a standalone window layout concern, though the layout structure is consumed by US-008-01, US-008-03, US-008-04, US-008-08)

## Acceptance Criteria
- [x] AC-01: The main window starts with a default size of 1280x720 pixels.
- [x] AC-02: The main window enforces a minimum size of 900x600 pixels that cannot be overridden at runtime.
- [x] AC-03: The main window is centered on screen using `WindowStartupLocation="CenterScreen"`.
- [x] AC-04: The main window uses a two-column Grid layout: fixed-width sidebar (250px left) and flexible content area (remaining width, `*`).
- [x] AC-05: The content area contains three rows: header (Auto height), content Frame (`*` fills remaining), and status bar (Auto height).
- [x] AC-06: The window background uses `BackgroundBrush` from the application resource dictionary.
- [x] AC-07: The sidebar width remains fixed at 250px while the content area fills the remaining space on resize.
- [x] AC-08: The content Frame expands to fill all available vertical space between the header and status bar.
- [x] AC-09: The window is resizable with proper content reflow and minimum size constraints enforced.

---

## TASK-008-05-01: Define MainWindow XAML root element with size, position, and background properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-008-05-02, TASK-008-05-03 |

### What to do
- Set `Height="720"` and `Width="1280"` on the `<Window>` element for default size
- Set `MinHeight="600"` and `MinWidth="900"` per VR-008-008 to enforce minimum constraints that cannot be overridden at runtime
- Set `WindowStartupLocation="CenterScreen"` for centered positioning
- Set `Background="{StaticResource BackgroundBrush}"` referencing the `#F5F5F5` resource from `App.xaml`
- Verify the skeleton already defines these attributes; confirm correctness

### How to verify
- [x] Window starts at 1280x720 (AC-01)
- [x] Window cannot be resized below 900x600 (AC-02)
- [x] Window is centered on screen (AC-03)
- [x] Background uses BackgroundBrush resource (AC-06)

---

## TASK-008-05-02: Define two-column Grid with fixed sidebar and star-sized content column

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | TASK-008-05-01 |
| Blocks | None |

### What to do
- Define the root `Grid` with two `ColumnDefinitions`: `<ColumnDefinition Width="250"/>` (sidebar, fixed) and `<ColumnDefinition Width="*"/>` (content, flexible)
- The sidebar column is fixed at 250px and does not change on window resize
- The content column uses `*` to fill all remaining horizontal space
- Verify the skeleton already has this structure; confirm correctness

### How to verify
- [x] Sidebar remains fixed at 250px on window resize (AC-04, AC-07)
- [x] Content area fills remaining horizontal space (AC-04, AC-07)
- [x] Layout reflows properly on resize (AC-09)

---

## TASK-008-05-03: Define three-row content area Grid

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | TASK-008-05-01 |
| Blocks | None |

### What to do
- Inside Grid column 1, define an inner `Grid` with three `RowDefinitions`: `<RowDefinition Height="Auto"/>` (header), `<RowDefinition Height="*"/>` (content Frame), `<RowDefinition Height="Auto"/>` (status bar)
- The header row auto-sizes to its content
- The content Frame row uses `*` to fill all remaining vertical space
- The status bar row auto-sizes to its content
- Verify the skeleton already has this structure; confirm correctness

### How to verify
- [x] Header row auto-sizes (AC-05)
- [x] Content Frame fills remaining vertical space between header and status bar (AC-05, AC-08)
- [x] Status bar row auto-sizes (AC-05)

---

## TASK-008-05-04: Define BackgroundBrush resource in application ResourceDictionary

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/App.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify `<Color x:Key="BackgroundColor">#F5F5F5</Color>` is defined in `App.xaml` ResourceDictionary
- Verify `<SolidColorBrush x:Key="BackgroundBrush" Color="{StaticResource BackgroundColor}"/>` is defined
- Both should already exist in the skeleton; confirm they are present and correctly valued
- Ensure `BackgroundBrush` is referenced by MainWindow via `{StaticResource BackgroundBrush}`

### How to verify
- [x] BackgroundBrush resource exists with value #F5F5F5 (AC-06)
- [x] MainWindow references BackgroundBrush, not a hardcoded color (AC-06)

---

## TASK-008-05-05: Set window Title and Icon

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/MainWindow.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Set `Title="Odoo AutoCAD Integration"` on the `<Window>` element
- Set `Icon` property to the application icon resource (e.g., `Icon="pack://application:,,,/Resources/odoo_autocad.ico"` or equivalent path)
- If the icon file is not yet included in the project resources, add it as an embedded resource or content file
- Wrap icon loading in try-catch if done in code-behind to handle missing icon gracefully

### How to verify
- [x] Window title bar displays "Odoo AutoCAD Integration" (AC-01)
- [x] Window icon is displayed in the title bar and taskbar (AC-01)

---

## Dependency Graph

```
TASK-008-05-01 (Window root attributes)
    ├── TASK-008-05-02 (Two-column Grid)
    └── TASK-008-05-03 (Three-row content Grid)

TASK-008-05-04 (BackgroundBrush resource) [independent]

TASK-008-05-05 (Title and Icon) [independent]
```
