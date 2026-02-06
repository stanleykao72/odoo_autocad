# US-008-02: Active Page Indicator

## User Story
**As a** CAD Engineer,
**I want to** see which page I am currently on via visual highlighting in the sidebar,
**So that** I always know my current location in the application.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The currently active navigation button has a visually distinct background color (e.g., `#E3F2FD`) differentiating it from inactive buttons.
- [ ] AC-02: The currently active navigation button displays a left border accent in the primary color (`PrimaryBrush`).
- [ ] AC-03: When navigating to a different page, the previously active button returns to its default (transparent) style.
- [ ] AC-04: The active button highlighting updates atomically with the Frame navigation to prevent desynchronization.
- [ ] AC-05: On application startup, the Dashboard button is highlighted as the active button.
- [ ] AC-06: Active state is driven by `MainViewModel.ActiveNavButton` property via data binding, not code-behind UI manipulation.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-010 | Active navigation button visually highlighted with distinct background and left border accent | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-02-01 | Add `ActiveNavButton` observable property to MainViewModel with default value "BtnDashboard" | `ViewModels/MainViewModel.cs` | S |
| TASK-008-02-02 | Implement NavButton style DataTriggers or value converter that compares button Tag/Name to ActiveNavButton | `Themes/Styles.xaml` | M |
| TASK-008-02-03 | Update NavigateCommand handler to set ActiveNavButton alongside CurrentPageTitle | `ViewModels/MainViewModel.cs` | S |
| TASK-008-02-04 | Assign Tag or Name property to each sidebar button matching the ActiveNavButton values | `Views/MainWindow.xaml` | S |

## Dependencies
- Depends on: US-008-01 (sidebar navigation buttons and NavButton style must exist)
- Blocks: None

## Notes
- The Python implementation does not highlight the active sidebar button since it uses direct action invocation rather than page navigation. This is a new behavior in the C# port.
- Three approaches were considered (see FR-008 Section 10): ViewModel binding, Style DataTrigger, and code-behind iteration. The recommended approach is ViewModel binding with DataTrigger for MVVM compliance.
- Validation rule VR-008-004 requires that active button highlighting updates atomically with Frame navigation to prevent visual desynchronization.
- The NavButton names map to pages: BtnDashboard, BtnAutoCAD, BtnOdoo, BtnBOQ, BtnPR, BtnMCP, BtnSettings.
