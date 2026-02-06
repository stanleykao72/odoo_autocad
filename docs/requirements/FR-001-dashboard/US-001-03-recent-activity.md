# US-001-03: Recent Activity

## User Story
**As a** Project Manager,
**I want to** see recent activity and sync status,
**So that** I know the current state of data flow between systems.

## Parent Feature
- **FR**: [FR-001-dashboard](../FR-001-dashboard/FR-001-dashboard.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Dashboard displays the last sync timestamp for Odoo data (e.g., "Last Sync: 2025-08-19 14:30")
- [ ] AC-02: Dashboard displays counts of recent operations including parameters extracted and BOQs pushed
- [ ] AC-03: Dashboard displays current project info (PR No, Project Name, Job Plan) when AutoCAD is connected and project context is available
- [ ] AC-04: When no project context is available, the project info panel displays "No project context available. Open a drawing with PR information."
- [ ] AC-05: Last sync timestamp updates automatically after each successful Odoo synchronization
- [ ] AC-06: Project info panel updates when a different AutoCAD drawing is opened or connection changes

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-001-010 | Dashboard SHOULD display last sync timestamp for Odoo data | Should |
| FR-001-011 | Dashboard SHOULD display count of recent operations (parameters extracted, BOQs pushed) | Could |
| FR-001-012 | Dashboard SHOULD display current project info (PR No, Project Name) when AutoCAD is connected | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-001-03-01 | Create Current Project info panel in XAML showing PR No, Project Name, Job Plan, and Last Sync fields bound to ViewModel properties | `Views/Pages/DashboardPage.xaml` | M |
| TASK-001-03-02 | Add ViewModel properties: CurrentPRNo, CurrentProjectName, CurrentJobPlanName, LastSyncTime (DateTime?) | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-03-03 | Implement logic to fetch current project context from IAutoCADService when connected and update ViewModel properties | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-03-04 | Implement logic to retrieve LastSyncTime from IOdooService.GetLastSyncTime() and display formatted timestamp | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-03-05 | Create Recent Operations summary section in XAML displaying counts of parameters extracted and BOQs pushed | `Views/Pages/DashboardPage.xaml` | M |
| TASK-001-03-06 | Add ViewModel properties and logic for tracking recent operation counts (ParametersExtractedCount, BOQsPushedCount) | `ViewModels/DashboardViewModel.cs` | M |
| TASK-001-03-07 | Handle "no project context" state: show placeholder message when AutoCAD is disconnected or no drawing with PR info is open | `ViewModels/DashboardViewModel.cs` | S |
| TASK-001-03-08 | Subscribe to AutoCAD document change events to refresh project info when a different drawing is opened | `ViewModels/DashboardViewModel.cs` | S |

## Dependencies
- Depends on: US-001-01 (connection status must be available to determine when to show project info)
- Depends on: IOdooService.GetLastSyncTime() being implemented
- Depends on: IAutoCADService providing current document and project context information
- Blocks: none

## Notes
- The project info panel in the wireframe shows: PR No, Project Name, Job Plan Name, and Last Sync time in a horizontal layout.
- LastSyncTime should be displayed in a user-friendly format (e.g., "2025-08-19 14:30") and handle null/never-synced state gracefully (e.g., "Never synced").
- FR-001-011 (operation counts) has priority "Could", so it can be implemented as a stretch goal; FR-001-010 and FR-001-012 with priority "Should" should be prioritized first.
- Operation counts could be persisted in the local SQLite database or held in-memory for the current session only, depending on requirements clarification.
- The Python implementation does not have a direct equivalent of the activity summary panel; this is a new feature for the C# version inspired by the richer dashboard card layout.
