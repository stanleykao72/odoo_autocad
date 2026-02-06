# US-008-07: Keyboard Shortcuts

## User Story
**As a** CAD Engineer,
**I want to** use keyboard shortcuts to navigate between pages,
**So that** I can work efficiently without relying solely on mouse interaction.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: `Ctrl+1` through `Ctrl+7` navigate to the seven sidebar pages in order: Dashboard, AutoCAD, Odoo, BOQ Manager, Purchase Requisition, AI Assistant, Settings.
- [ ] AC-02: `F5` triggers the Refresh command (equivalent to clicking the Refresh button in the header).
- [ ] AC-03: `Alt+Left` navigates back to the previous page when `CanGoBack` is true in the navigation journal.
- [ ] AC-04: Keyboard shortcuts are registered as `InputBinding` or `KeyBinding` entries in MainWindow XAML.
- [ ] AC-05: Keyboard shortcuts work regardless of which page is currently displayed.
- [ ] AC-06: When `Alt+Left` is pressed and `CanGoBack` is false, the shortcut is silently ignored (no error or feedback).

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-042 | Ctrl+1 through Ctrl+7 for navigating to the seven sidebar pages | Should |
| FR-008-043 | F5 as shortcut for Refresh command | Should |
| FR-008-044 | Alt+Left for navigating back when CanGoBack is true | Could |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-07-01 | Define InputBindings in MainWindow.xaml for Ctrl+1 through Ctrl+7 mapped to NavigateCommand with page parameter | `Views/MainWindow.xaml` | M |
| TASK-008-07-02 | Define InputBinding for F5 mapped to RefreshCommand | `Views/MainWindow.xaml` | S |
| TASK-008-07-03 | Define InputBinding for Alt+Left mapped to GoBackCommand | `Views/MainWindow.xaml` | S |
| TASK-008-07-04 | Add GoBackCommand to MainViewModel that delegates to INavigationService.GoBack() with CanGoBack guard | `ViewModels/MainViewModel.cs` | S |
| TASK-008-07-05 | Ensure NavigateCommand accepts a string parameter for page name routing from keyboard shortcuts | `ViewModels/MainViewModel.cs` | S |

## Dependencies
- Depends on: US-008-01 (NavigateCommand and INavigationService must exist)
- Blocks: None

## Notes
- The Python implementation does not have keyboard shortcuts for navigation. This is a new feature in the C# port to improve workflow efficiency.
- WPF `InputBinding` / `KeyBinding` elements in the Window's `InputBindings` collection provide the cleanest MVVM-compatible approach.
- The Ctrl+N shortcuts map by index: Ctrl+1 = Dashboard, Ctrl+2 = AutoCAD, Ctrl+3 = Odoo, Ctrl+4 = BOQ Manager, Ctrl+5 = Purchase Requisition, Ctrl+6 = AI Assistant, Ctrl+7 = Settings.
- Validation rule VR-008-002 requires `GoBack()` to check `CanGoBack` before invoking `Frame.GoBack()` to avoid `InvalidOperationException`.
