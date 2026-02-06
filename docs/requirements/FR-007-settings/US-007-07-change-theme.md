# US-007-07: Change Theme

## User Story
**As a** CAD Engineer,
**I want to** change the application theme (light, dark, system),
**So that** the UI matches my preference and reduces eye strain.

## Parent Feature
- **FR**: [FR-007-settings](../FR-007-settings/FR-007-settings.md)
- **Priority**: P2

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

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-007-013 | Settings Page SHALL provide a theme selector with options: System, Light, Dark | Must |
| FR-007-014 | Theme changes SHALL be applied immediately as a live preview without requiring application restart | Should |
| FR-007-026 | Settings Page SHALL provide a Save button that persists all modified settings | Must |
| FR-007-028 | All settings changes SHALL be logged to the SyncLog table | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-007-07-01 | Add theme radio buttons (System, Light, Dark) to Appearance tab in XAML with StringMatchConverter binding | `Views/Pages/SettingsPage.xaml` | S |
| TASK-007-07-02 | Add SelectedTheme property to ViewModel that triggers live preview on change | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-07-03 | Implement IThemeService.ApplyTheme() to swap WPF ResourceDictionary for live preview | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-07-04 | Create Light and Dark theme ResourceDictionary XAML files with appropriate color schemes | `Views/Pages/SettingsPage.xaml` | M |
| TASK-007-07-05 | Implement "System" theme detection that follows Windows OS light/dark mode setting | `ViewModels/SettingsViewModel.cs` | M |
| TASK-007-07-06 | Persist theme selection to UserPreferences table on Save | `ViewModels/SettingsViewModel.cs` | S |
| TASK-007-07-07 | Load theme preference on application startup and apply before main window is shown | `ViewModels/SettingsViewModel.cs` | S |

## Dependencies
- Depends on: None (independent settings section)
- Blocks: None

## Notes
- The Python application uses `CustomTkinter.set_appearance_mode()` with modes "system", "light", "dark" and color themes "blue", "green", "dark-blue". The C# migration uses WPF ResourceDictionary merging via `IThemeService`.
- The primary color from Python is `#2B579A` (Professional Blue). This should be carried forward to the C# theme dictionaries.
- Validation rule: VR-007-010 (theme must be one of "System", "Light", "Dark").
- ConfigKey mapping: `appearance.theme` in `UserPreferences` table.
- The SelectedTheme property setter should directly call `IThemeService.ApplyTheme()` for immediate preview, but the value is only persisted to the database when the Save button is clicked.
