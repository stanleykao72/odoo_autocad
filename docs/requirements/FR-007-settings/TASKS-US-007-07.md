# TASKS: US-007-07 — Change Theme

> **Parent US**: [US-007-07](US-007-07-change-theme.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 4S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] `AppDbContext` with `UserPreferences` and `SyncLogs` DbSets available
- [ ] `IThemeService` interface defined (or to be created as part of this US)

## Acceptance Criteria
- [ ] AC-01: Settings Page displays theme selector with three radio button options: System, Light, Dark in the Appearance tab
- [ ] AC-02: The currently active theme is pre-selected on page load
- [ ] AC-03: Selecting a theme applies the change immediately as a live preview without requiring application restart
- [ ] AC-04: Live preview swaps the WPF ResourceDictionary (merging Light or Dark theme dictionary) via `IThemeService.ApplyTheme()`
- [ ] AC-05: Theme selection must be one of: "System", "Light", "Dark" (validated)
- [ ] AC-06: Saving persists the theme selection to `UserPreferences` table with key `appearance.theme`
- [ ] AC-07: Theme preference is loaded from `UserPreferences` table on application startup and applied before the main window is shown
- [ ] AC-08: "System" theme follows the Windows OS theme setting (light/dark mode)
- [ ] AC-09: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-07-01: Add theme radio buttons (System, Light, Dark) to Appearance tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Populate the Appearance tab (stub created in US-007-01) with a "Theme" section
- Add three `RadioButton` controls in a vertical `StackPanel`:
  - `Content="System"` with `IsChecked="{Binding SelectedTheme, Converter={StaticResource StringMatchConverter}, ConverterParameter=System}"`
  - `Content="Light"` with `IsChecked="{Binding SelectedTheme, Converter={StaticResource StringMatchConverter}, ConverterParameter=Light}"`
  - `Content="Dark"` with `IsChecked="{Binding SelectedTheme, Converter={StaticResource StringMatchConverter}, ConverterParameter=Dark}"`
- Create a `StringMatchConverter` (if not already existing) in `Converters/StringMatchConverter.cs` that returns true when the bound value equals the `ConverterParameter`
- Group the radio buttons with `GroupName="ThemeGroup"` to ensure mutual exclusivity

### How to verify
- [ ] Three radio buttons render in the Appearance tab: System, Light, Dark (AC-01)
- [ ] Only one radio button can be selected at a time (AC-01)

---

## TASK-007-07-02: Add SelectedTheme property to ViewModel that triggers live preview on change

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-07-03, TASK-007-07-06 |

### What to do
- Add `[ObservableProperty] string _selectedTheme` with default `"System"`
- In `OnSelectedThemeChanged`, call `_themeService.ApplyTheme(SelectedTheme)` for immediate live preview
- Wire `OnSelectedThemeChanged` to also update `HasUnsavedChanges`
- Inject `IThemeService` via constructor (add to DI registration)

### How to verify
- [ ] Changing the radio button selection immediately updates `SelectedTheme` (AC-01)
- [ ] Theme change triggers `IThemeService.ApplyTheme()` (AC-03)

---

## TASK-007-07-03: Implement IThemeService with ApplyTheme() to swap WPF ResourceDictionary

| Field | Value |
|-------|-------|
| Target | `Services/IThemeService.cs` and `Services/ThemeService.cs` |
| Estimate | M |
| Depends On | TASK-007-07-04 |
| Blocks | TASK-007-07-02 |

### What to do
- Create `IThemeService` interface with method `void ApplyTheme(string themeName)` where themeName is "System", "Light", or "Dark"
- Implement `ThemeService`:
  - `ApplyTheme("Light")`: clear existing theme dictionaries from `Application.Current.Resources.MergedDictionaries`, then add `LightTheme.xaml` ResourceDictionary
  - `ApplyTheme("Dark")`: clear and add `DarkTheme.xaml` ResourceDictionary
  - `ApplyTheme("System")`: detect Windows OS theme (see TASK-007-07-05), then apply Light or Dark accordingly
- The method must run on the UI thread; use `Application.Current.Dispatcher.Invoke()` if called from a non-UI context
- Register `ThemeService` as singleton in DI since the theme state is global

### How to verify
- [ ] `ApplyTheme("Light")` applies the light theme ResourceDictionary immediately (AC-03, AC-04)
- [ ] `ApplyTheme("Dark")` applies the dark theme ResourceDictionary immediately (AC-03, AC-04)
- [ ] Theme switch does not require application restart (AC-03)

---

## TASK-007-07-04: Create Light and Dark theme ResourceDictionary XAML files

| Field | Value |
|-------|-------|
| Target | `Resources/Themes/LightTheme.xaml` and `Resources/Themes/DarkTheme.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-07-03 |

### What to do
- Create `Resources/Themes/LightTheme.xaml` with `ResourceDictionary` defining color brushes:
  - `BackgroundBrush`: White/light gray
  - `SurfaceBrush`: White
  - `TextPrimaryBrush`: Dark gray/black
  - `TextSecondaryBrush`: Medium gray
  - `AccentBrush`: `#2B579A` (Professional Blue, carried from Python theme)
  - `BorderBrush`: Light gray
- Create `Resources/Themes/DarkTheme.xaml` with corresponding dark variants:
  - `BackgroundBrush`: Dark gray (`#1E1E1E`)
  - `SurfaceBrush`: Slightly lighter gray (`#252526`)
  - `TextPrimaryBrush`: White/light gray
  - `TextSecondaryBrush`: Medium gray
  - `AccentBrush`: `#4A90D9` (Lighter blue for dark backgrounds)
  - `BorderBrush`: Dark border (`#3C3C3C`)
- Define consistent control styles (Button, TextBox, ComboBox, TabControl, etc.) in each theme
- Ensure both dictionaries define the same set of resource keys so they are interchangeable

### How to verify
- [ ] LightTheme.xaml and DarkTheme.xaml compile without errors (AC-04)
- [ ] Both files define the same resource keys (AC-04)
- [ ] Primary accent color is `#2B579A` (carried from Python) (AC-04)

---

## TASK-007-07-05: Implement "System" theme detection following Windows OS light/dark mode

| Field | Value |
|-------|-------|
| Target | `Services/ThemeService.cs` |
| Estimate | M |
| Depends On | TASK-007-07-03 |
| Blocks | None |

### What to do
- Read the Windows registry key `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme` to detect the OS preference:
  - Value `1` = Light mode -> apply Light theme
  - Value `0` = Dark mode -> apply Dark theme
- Alternatively, use `Microsoft.Win32.SystemParameters` or the Windows Runtime `UISettings` API for theme detection
- Subscribe to the `SystemEvents.UserPreferenceChanged` event to detect live OS theme changes and re-apply the "System" theme accordingly
- Fall back to Light theme if OS detection fails

### How to verify
- [ ] When "System" is selected and OS is in dark mode, the dark theme is applied (AC-08)
- [ ] When "System" is selected and OS is in light mode, the light theme is applied (AC-08)
- [ ] Changing OS theme while "System" is selected updates the app theme automatically (AC-08)

---

## TASK-007-07-06: Persist theme selection to UserPreferences table on Save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-07-02 |
| Blocks | TASK-007-07-07 |

### What to do
- In the `SaveSettingsCommand` handler, add persistence for theme:
  - Upsert `UserPreferences` row: `PreferenceKey = "appearance.theme"`, `PreferenceValue = SelectedTheme`, `DataType = "string"`
- Include the theme key in the `SyncLog` details JSON
- Validate that `SelectedTheme` is one of "System", "Light", "Dark" before persisting (VR-007-010)
- Update the original-values snapshot for `SelectedTheme` after save

### How to verify
- [ ] After Save, `UserPreferences` table contains `appearance.theme` with the selected value (AC-06)
- [ ] `SyncLog` entry includes the theme change (AC-09)
- [ ] Invalid theme values are rejected (AC-05)

---

## TASK-007-07-07: Load theme preference on application startup and apply before main window is shown

| Field | Value |
|-------|-------|
| Target | `App.xaml.cs` |
| Estimate | S |
| Depends On | TASK-007-07-06 |
| Blocks | None |

### What to do
- In `App.OnStartup()` or `ConfigureServices()`, after DI container is built:
  1. Query `UserPreferences` for `appearance.theme`; fall back to "System"
  2. Resolve `IThemeService` from the DI container
  3. Call `themeService.ApplyTheme(savedTheme)` **before** creating and showing `MainWindow`
- This ensures the correct theme is applied from the first frame, avoiding a flash of default theme
- In the SettingsViewModel constructor, load the theme preference and set `SelectedTheme` to match

### How to verify
- [ ] Application starts with the previously saved theme (no flash of default) (AC-07)
- [ ] On first launch (no saved preference), "System" theme is used (AC-07, AC-08)

---

## Dependency Graph
```
TASK-007-07-04 (Theme ResourceDictionaries)
       │
       └──▶ TASK-007-07-03 (IThemeService Implementation) ──▶ TASK-007-07-02 (ViewModel Property)
                  │
                  └──▶ TASK-007-07-05 (System Theme Detection)

TASK-007-07-01 (XAML Radio Buttons)  ── independent

TASK-007-07-02 (ViewModel Property)
       │
       └──▶ TASK-007-07-06 (Persist on Save)
                  │
                  └──▶ TASK-007-07-07 (Startup Load)

Task 04 and 01 can start in parallel.
Task 03 depends on 04 (theme files must exist).
Task 02 depends on 03 (needs IThemeService).
```
