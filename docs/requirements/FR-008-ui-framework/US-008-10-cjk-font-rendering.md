# US-008-10: CJK Font Rendering

## User Story
**As a** CAD Engineer,
**I want to** have CJK characters render correctly throughout the application,
**So that** all Chinese labels, status messages, and log entries are legible.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The application uses Microsoft JhengHei UI as the primary font family for all UI elements.
- [ ] AC-02: A font fallback chain is defined: Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS.
- [ ] AC-03: The font family is set at the Window or Application level so it applies globally without per-element configuration.
- [ ] AC-04: All Chinese labels, status messages, log entries, and user-facing text render correctly without missing glyphs or tofu characters.
- [ ] AC-05: The font family resource is defined as a named resource (`AppFontFamily`) in the resource dictionary for consistent reference.
- [ ] AC-06: All styles and templates reference the `AppFontFamily` resource rather than hardcoding font names.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-035 | CJK font rendering using Microsoft JhengHei UI with appropriate fallbacks | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-10-01 | Define AppFontFamily resource with fallback chain in Fonts resource dictionary | `Themes/Fonts.xaml` | S |
| TASK-008-10-02 | Set FontFamily at the Window level in MainWindow.xaml referencing AppFontFamily | `Views/MainWindow.xaml` | S |
| TASK-008-10-03 | Merge Fonts.xaml into App.xaml resource dictionary | `App.xaml` | S |
| TASK-008-10-04 | Verify all named styles (NavButton, PrimaryButton, CardPanel, ModernTextBox) reference AppFontFamily | `Themes/Styles.xaml` | S |
| TASK-008-10-05 | Test CJK rendering with sample Chinese strings across all UI areas (sidebar, header, status bar, log panel) | `Tests/` | M |

## Dependencies
- Depends on: None (font configuration is independent)
- Blocks: None

## Notes
- The Python implementation uses `FontManager` with a cross-platform fallback chain. On Windows, the priority is: Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS. The C# port uses the same chain defined as a WPF `FontFamily` resource.
- Python font sizes: title=16pt, heading=14pt, body=11pt, small=9pt, button=10pt. The C# port should define equivalent font size resources for consistency.
- WPF handles CJK fallback natively through composite font families. Setting the `FontFamily` at the `Window` level with a comma-separated list provides automatic fallback.
- The recommended approach is:
  ```xml
  <FontFamily x:Key="AppFontFamily">Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS</FontFamily>
  ```
- This is a cross-cutting concern referenced by FR-008-031 through FR-008-034 (theme/styling requirements). All styles must use `{StaticResource AppFontFamily}` rather than hardcoded font names.
