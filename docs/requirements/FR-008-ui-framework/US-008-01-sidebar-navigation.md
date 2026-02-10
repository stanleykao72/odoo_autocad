# US-008-01: Sidebar Navigation

## User Story
**As a** CAD Engineer,
**I want to** navigate between different functional areas using a sidebar menu,
**So that** I can quickly access AutoCAD, Odoo, BOQ, and PR features without losing context.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [x] AC-01: A fixed-width (250px) sidebar panel is displayed on the left side of the main window at all times.
- [x] AC-02: The sidebar displays a title area at the top with "Odoo AutoCAD" and "Integration System v6.0".
- [x] AC-03: The sidebar contains navigation buttons for all seven areas: Dashboard, AutoCAD, Odoo, BOQ Manager, Purchase Requisition, AI Assistant, Settings.
- [x] AC-04: Each navigation button displays an icon and a text label, left-aligned.
- [x] AC-05: Navigation buttons use the `NavButton` style with transparent background, left-aligned content, and a left-border accent on hover.
- [x] AC-06: Clicking a sidebar button loads the corresponding page into the main `Frame` via `INavigationService.NavigateTo()`.
- [x] AC-07: The application defaults to the Dashboard page on startup.
- [x] AC-08: The `Frame` element uses `NavigationUIVisibility="Hidden"` to suppress default WPF navigation chrome.
- [x] AC-09: `INavigationService.NavigateTo(pageName)` resolves pages by reflection from namespace `OdooAutoCAD.App.Views.Pages.{pageName}Page`.
- [x] AC-10: `INavigationService.GoBack()` navigates to the previous page when `CanGoBack` is true.
- [x] AC-11: If a resolved page type does not exist, navigation is silently skipped and a warning is logged.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-006 | Sidebar title area with application name and version | Must |
| FR-008-007 | Navigation buttons for all seven functional areas | Must |
| FR-008-008 | Icon and text label on each navigation button, left-aligned | Must |
| FR-008-009 | Consistent NavButton style with hover accent | Must |
| FR-008-010 | Active navigation button visual highlighting | Must |
| FR-008-011 | Connection Status section at bottom of sidebar | Must |
| FR-008-012 | Frame element with hidden navigation UI for page hosting | Must |
| FR-008-013 | Page resolution by reflection from namespace | Must |
| FR-008-014 | GoBack navigation via Frame journal | Should |
| FR-008-015 | Default navigation to Dashboard on startup | Must |
| FR-008-016 | Update CurrentPageTitle on navigation | Must |
| FR-008-017 | Update sidebar button highlighting on navigation | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-01-01 | Define sidebar XAML layout with title area, navigation buttons, and connection status section | `Views/MainWindow.xaml` | L |
| TASK-008-01-02 | Implement NavButton style with transparent background, left-align, and hover left-border accent | `Themes/Styles.xaml` | M |
| TASK-008-01-03 | Create INavigationService interface with NavigateTo, GoBack, CanGoBack, and Frame property | `Services/NavigationService.cs` | M |
| TASK-008-01-04 | Implement NavigationService with reflection-based page resolution from OdooAutoCAD.App.Views.Pages namespace | `Services/NavigationService.cs` | M |
| TASK-008-01-05 | Add NavigateCommand and page title mapping to MainViewModel | `ViewModels/MainViewModel.cs` | M |
| TASK-008-01-06 | Wire MainWindow constructor to set Frame on NavigationService and navigate to Dashboard | `Views/MainWindow.xaml.cs` | S |
| TASK-008-01-07 | Register INavigationService as singleton in DI container | `App.xaml.cs` | S |

## Dependencies
- Depends on: None (this is the foundational US for the entire UI framework)
- Blocks: US-008-02, US-008-03, US-008-04, US-008-07, US-008-08, and all feature-area pages across other FRs

## Notes
- This is the foundational user story for the entire application shell. All other UI-related user stories in FR-008 and other feature FRs depend on the sidebar and navigation infrastructure established here.
- The Python implementation uses dynamic widget replacement (`clear_main_content()`), whereas the C# port uses WPF `Frame`-based navigation with a proper page journal.
- Page types are resolved via `Type.GetType($"OdooAutoCAD.App.Views.Pages.{pageName}Page")`. Each feature area must create a corresponding `Page` class in that namespace.
- The sidebar sections in Python are: "連接管理" (Connect), "主要功能" (Main), "工具" (Tools), and SSE Control. The C# port consolidates these into seven clean navigation buttons plus a connection status panel.
- Validation rule VR-008-003 requires that `INavigationService.Frame` is set before any navigation calls; the MainWindow constructor must assign `MainFrame` to the service.
