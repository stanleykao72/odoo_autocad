# US-008-08: Page Title Header

## User Story
**As a** CAD Engineer,
**I want to** see the current page title in the header area,
**So that** I have clear confirmation of what page is displayed.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A header bar is displayed at the top of the content area (above the main Frame, to the right of the sidebar).
- [ ] AC-02: The header bar displays `CurrentPageTitle` as a 20pt SemiBold TextBlock bound to the MainViewModel.
- [ ] AC-03: The header bar includes a "Refresh" button on the right side bound to `RefreshCommand`.
- [ ] AC-04: The header bar optionally includes a "Help" button on the right side.
- [ ] AC-05: The header bar uses `SurfaceBrush` background with a bottom border separator.
- [ ] AC-06: `CurrentPageTitle` updates automatically when navigation occurs, reflecting the display name of the active page.
- [ ] AC-07: Page title mapping follows the defined convention: Dashboard, AutoCAD Integration, Odoo Connection, BOQ Manager, Purchase Requisition, AI Assistant (MCP), Settings.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-018 | Header displays CurrentPageTitle as 20pt SemiBold TextBlock | Must |
| FR-008-019 | Header includes Refresh button bound to RefreshCommand | Should |
| FR-008-020 | Header includes Help button | Could |
| FR-008-021 | Header uses SurfaceBrush background with bottom border | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-08-01 | Define header bar XAML with page title TextBlock (20pt SemiBold) and action buttons (Refresh, Help) | `Views/MainWindow.xaml` | M |
| TASK-008-08-02 | Add CurrentPageTitle observable property to MainViewModel with default "Dashboard" | `ViewModels/MainViewModel.cs` | S |
| TASK-008-08-03 | Add RefreshCommand to MainViewModel that triggers page-specific refresh logic | `ViewModels/MainViewModel.cs` | S |
| TASK-008-08-04 | Define page title display name mapping (NavButton name to display title) in NavigationService or MainViewModel | `ViewModels/MainViewModel.cs` | S |
| TASK-008-08-05 | Apply SurfaceBrush background and bottom border styling to header bar | `Themes/Styles.xaml` | S |

## Dependencies
- Depends on: US-008-01 (navigation must update CurrentPageTitle on page change)
- Blocks: None

## Notes
- The Python implementation uses a top bar with "AutoCAD Odoo 整合系統" as a fixed title with primary color background and white text. The C# port replaces this with a context-sensitive page title header that updates on navigation.
- The page title mapping is defined in FR-008 Section 6: BtnDashboard -> "Dashboard", BtnAutoCAD -> "AutoCAD Integration", BtnOdoo -> "Odoo Connection", BtnBOQ -> "BOQ Manager", BtnPR -> "Purchase Requisition", BtnMCP -> "AI Assistant (MCP)", BtnSettings -> "Settings".
- FR-008-016 requires that navigation updates `MainViewModel.CurrentPageTitle` to reflect the display name of the active page. This links US-008-08 directly to the navigation logic in US-008-01.
