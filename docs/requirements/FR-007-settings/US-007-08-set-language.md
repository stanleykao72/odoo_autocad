# US-007-08: Set Language

## User Story
**As a** CAD Engineer,
**I want to** set the application language/locale,
**So that** I can use the application in my preferred language.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a language/locale selector dropdown in the Appearance tab with options: English (en-US), Traditional Chinese (zh-TW)
- [ ] AC-02: The currently active language is pre-selected on page load
- [ ] AC-03: Language selection must be one of the supported locale codes: "en-US", "zh-TW" (validated)
- [ ] AC-04: A notification message "Language change requires restart" is displayed below the selector when a different language is chosen
- [ ] AC-05: Language changes take effect after application restart, not immediately
- [ ] AC-06: Saving persists the language selection to `UserPreferences` table with key `appearance.language`
- [ ] AC-07: Language preference is loaded from `UserPreferences` table on application startup and applied to the UI culture
- [ ] AC-08: Settings change is logged to the `SyncLog` table with SyncType "settings_change"

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-015 | Settings Page SHALL provide a language/locale selector (English, Traditional Chinese) | Should |
| FR-007-016 | Language changes SHALL take effect after application restart, with a notification informing the user | Should |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-08-01 | Add language selector dropdown to Appearance tab in XAML bound to SelectedLanguage | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-08-02 | Add restart notification text element that appears when language differs from current active language | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-08-03 | Add SelectedLanguage property to ViewModel with supported locale codes | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-08-04 | Persist language selection to UserPreferences table on Save | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-08-05 | Load language preference on application startup and set CultureInfo/UICulture before main window is created | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-08-06 | Create resource files (.resx) for English and Traditional Chinese string localization | `Views/Pages/SettingsPage.xaml` | L |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The Python application has hardcoded Traditional Chinese (zh-TW) support using Microsoft JhengHei UI font. The C# migration adds proper localization infrastructure with .resx resource files.
- Language change requires restart because WPF resource strings are loaded at XAML parse time. The notification should clearly inform the user that a restart is needed.
- Validation rule: VR-007-011 (language must be one of "en-US", "zh-TW").
- ConfigKey mapping: `appearance.language` in `UserPreferences` table.
- On startup, the application reads the language preference and sets `Thread.CurrentThread.CurrentCulture` and `Thread.CurrentThread.CurrentUICulture` before creating the main window.
- Font consideration: When zh-TW is selected, ensure Microsoft JhengHei UI is used as the primary font family (matching the Python application behavior).
