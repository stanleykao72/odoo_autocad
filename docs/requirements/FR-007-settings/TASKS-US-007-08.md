# TASKS: US-007-08 — Set Language

> **Parent US**: [US-007-08](US-007-08-set-language.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 3S + 2M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] US-007-07 (Appearance tab must exist for language selector to be placed in)
- [ ] `AppDbContext` with `UserPreferences` and `SyncLogs` DbSets available

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a language/locale selector dropdown in the Appearance tab with options: English (en-US), Traditional Chinese (zh-TW)
- [ ] AC-02: The currently active language is pre-selected on page load
- [ ] AC-03: Language selection must be one of the supported locale codes: "en-US", "zh-TW" (validated)
- [ ] AC-04: A notification message "Language change requires restart" is displayed below the selector when a different language is chosen
- [ ] AC-05: Language changes take effect after application restart, not immediately
- [ ] AC-06: Saving persists the language selection to `UserPreferences` table with key `appearance.language`
- [ ] AC-07: Language preference is loaded from `UserPreferences` table on application startup and applied to the UI culture
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-08-01: Add language selector dropdown to Appearance tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Appearance tab (below the theme section from US-007-07), add a "Language" section
- Add a `ComboBox` with label "Language:" bound to `{Binding SelectedLanguage}`
- Define `ComboBoxItem` entries: `Content="English"` with `Tag="en-US"` and `Content="Traditional Chinese (繁體中文)"` with `Tag="zh-TW"`
- Alternatively, bind `ItemsSource` to a collection of language display objects with `DisplayMember` and `ValueMember` properties
- Style consistently with other Appearance tab controls

### How to verify
- [ ] Language dropdown renders with English and Traditional Chinese options (AC-01)
- [ ] Dropdown is positioned in the Appearance tab below the theme selector (AC-01)

---

## TASK-007-08-02: Add restart notification text that appears when language differs from current active language

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `TextBlock` below the language dropdown with text "* Language change requires restart"
- Bind `Visibility` to a computed property (e.g., `IsLanguageChangeNotificationVisible`) that returns `Visible` when `SelectedLanguage` differs from the originally loaded language, and `Collapsed` otherwise
- Style with italic text and a warning color (e.g., orange or muted text)
- The notification should not block any action -- it is purely informational

### How to verify
- [ ] Notification text appears when a different language is selected (AC-04)
- [ ] Notification text is hidden when the language matches the current active language (AC-04)

---

## TASK-007-08-03: Add SelectedLanguage property to ViewModel with supported locale codes

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-08-04, TASK-007-08-06 |

### What to do
- Add `[ObservableProperty] string _selectedLanguage` with default `"zh-TW"` (matching the Python application's default)
- Add a `string _originalLanguage` field to track the language at page load (for restart notification)
- Add a computed property `bool IsLanguageChangeNotificationVisible` that returns `SelectedLanguage != _originalLanguage`
- In `OnSelectedLanguageChanged`, call `OnPropertyChanged(nameof(IsLanguageChangeNotificationVisible))` and update `HasUnsavedChanges`
- Validate that `SelectedLanguage` is one of `"en-US"` or `"zh-TW"` (VR-007-011)

### How to verify
- [ ] `SelectedLanguage` property is bindable and defaults to "zh-TW" (AC-01, AC-02)
- [ ] `IsLanguageChangeNotificationVisible` correctly reflects whether a restart is needed (AC-04)
- [ ] Invalid locale codes are rejected (AC-03)

---

## TASK-007-08-04: Persist language selection to UserPreferences table on Save

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-08-03 |
| Blocks | TASK-007-08-05 |

### What to do
- In the `SaveSettingsCommand` handler, add persistence for language:
  - Upsert `UserPreferences` row: `PreferenceKey = "appearance.language"`, `PreferenceValue = SelectedLanguage`, `DataType = "string"`
- Include the language key in the `SyncLog` details JSON with `SyncType = "settings_change"`
- Update the original-values snapshot for `SelectedLanguage` after save (but note: `_originalLanguage` for the restart notification should NOT be updated, since the change still requires restart)

### How to verify
- [ ] After Save, `UserPreferences` table contains `appearance.language` with selected locale code (AC-06)
- [ ] `SyncLog` entry includes the language change (AC-08)

---

## TASK-007-08-05: Load language preference on application startup and set CultureInfo before main window

| Field | Value |
|-------|-------|
| Target | `App.xaml.cs` |
| Estimate | M |
| Depends On | TASK-007-08-04 |
| Blocks | None |

### What to do
- In `App.OnStartup()`, after building the DI container but **before** creating `MainWindow`:
  1. Query `UserPreferences` for `appearance.language`; fall back to `"zh-TW"`
  2. Create a `CultureInfo` from the locale code: `var culture = new CultureInfo(savedLanguage)`
  3. Set `Thread.CurrentThread.CurrentCulture = culture` and `Thread.CurrentThread.CurrentUICulture = culture`
  4. Optionally set `FrameworkElement.LanguageProperty.OverrideMetadata(typeof(FrameworkElement), new FrameworkPropertyMetadata(XmlLanguage.GetLanguage(culture.IetfLanguageTag)))` to propagate to all WPF elements
- When `zh-TW` is selected, configure the default font family to use `Microsoft JhengHei UI` (matching Python behavior)
- Log the applied language via `ILogger`

### How to verify
- [ ] Application starts with the previously saved language/culture (AC-07)
- [ ] WPF elements use the correct culture for formatting (dates, numbers) (AC-05, AC-07)
- [ ] On first launch, default is "zh-TW" with Microsoft JhengHei UI font (AC-07)

---

## TASK-007-08-06: Create resource files (.resx) for English and Traditional Chinese string localization

| Field | Value |
|-------|-------|
| Target | `Resources/Strings/Strings.resx` and `Resources/Strings/Strings.zh-TW.resx` |
| Estimate | L |
| Depends On | TASK-007-08-03 |
| Blocks | None |

### What to do
- Create `Resources/Strings/Strings.resx` (default/English) with all user-facing strings in the application:
  - Settings page labels: "Server URL", "Database Name", "Username", "API Token", "Connection Timeout", "Test Connection", "Save Settings", "Cancel", etc.
  - Tab names: "Connection", "AutoCAD", "MCP", "Appearance", "Advanced"
  - Error messages: validation errors, connection test results, dialog messages
  - Common strings: "Connected", "Disconnected", "Loading...", "Success", "Error", etc.
- Create `Resources/Strings/Strings.zh-TW.resx` with Traditional Chinese translations for all keys:
  - "Server URL" -> "伺服器網址", "Database Name" -> "資料庫名稱", "Save Settings" -> "儲存設定", etc.
- Update XAML bindings to use `x:Static` references to the generated `Strings` class: `Text="{x:Static strings:Strings.ServerUrl}"`
- Configure the `.resx` files to generate a public `Strings` class with `ResXFileCodeGenerator` custom tool
- This is a large task as it involves extracting and translating all hardcoded strings across the application

### How to verify
- [ ] All user-facing strings are sourced from resource files, not hardcoded (AC-05)
- [ ] Switching to English shows English strings; switching to zh-TW shows Chinese strings after restart (AC-01, AC-05)
- [ ] No missing translations (all keys present in both .resx files) (AC-01)

---

## Dependency Graph
```
TASK-007-08-01 (XAML Dropdown)       ── independent
TASK-007-08-02 (Restart Notification) ── independent

TASK-007-08-03 (ViewModel Property)
       │
       ├──▶ TASK-007-08-04 (Persist on Save)
       │         │
       │         └──▶ TASK-007-08-05 (Startup Culture)
       │
       └──▶ TASK-007-08-06 (Resource Files)

Tasks 01, 02, and 03 can start in parallel.
Task 06 is the largest and can start once 03 defines the language codes.
```
