# US-003-07: Search Projects

## User Story
**As a** CAD Engineer,
**I want to** search for a project by PR number or name,
**So that** I can associate AutoCAD drawings with the correct Odoo project.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A project search field is visible on the Odoo Integration Page
- [ ] AC-02: The search queries Odoo by project name or code via `SearchProjectsAsync(searchTerm)`
- [ ] AC-03: Search results display in a list/DataGrid with columns: ID, Name, Code, State, Date Start, Date End
- [ ] AC-04: The user can select a project from search results to set it as the active project context
- [ ] AC-05: The selected project is available for downstream operations (BOQ import, PR generation)
- [ ] AC-06: The search field is disabled when `IsConnected` is false
- [ ] AC-07: When no projects match, the message "No projects found matching '{searchTerm}'. Try a different search term." is displayed
- [ ] AC-08: Single project retrieval by ID is supported via `GetProjectAsync(projectId)` for navigation from other pages
- [ ] AC-09: The search operation runs asynchronously without blocking the UI thread

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-017 | The page SHALL provide a project search field that queries Odoo by project name or code via `SearchProjectsAsync(searchTerm)` | Must |
| FR-003-018 | Project search results SHALL display in a list with columns: ID, Name, Code, State, Date Start, Date End | Must |
| FR-003-019 | The user SHALL be able to select a project from search results to set it as the active project context for downstream operations | Must |
| FR-003-020 | The page SHALL support retrieving a single project by ID via `GetProjectAsync(projectId)` when navigating from other pages | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-07-01 | Create Project Search panel in XAML with TextBox, Search button, and results DataGrid | `Views/Pages/OdooConnectionPage.xaml` | M |
| TASK-003-07-02 | Add ProjectSearchTerm, ProjectSearchResults, and SelectedProject properties to ViewModel | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-07-03 | Implement SearchProjectsCommand as IAsyncRelayCommand calling IOdooService.SearchProjectsAsync(searchTerm) | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-07-04 | Implement SelectProjectCommand as IRelayCommand that sets SelectedProject as active context | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-07-05 | Implement GetProjectAsync(projectId) in OdooService for single-project retrieval from other pages | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-07-06 | Bind search field IsEnabled to IsConnected and add empty state message for no results | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-07-07 | Add "Select Project" button below results DataGrid to confirm project selection | `Views/Pages/OdooConnectionPage.xaml` | S |

## Dependencies
- Depends on: US-003-01, US-003-03
- Blocks: None

## Notes
- The Python equivalent is `get_project(pr_no)` which calls `get_project_v2` with domain `[['name', '=', pr_no]]`.
- The C# version uses `SearchProjectsAsync(searchTerm)` for flexible search by name or code, and `GetProjectAsync(projectId)` for direct lookup.
- Validation rule VR-003-009 ensures project search is only available when `IsConnected` is true.
- Validation rule VR-003-011 requires a selected project for BOQ sync operations.
- The selected project context is critical for downstream workflows: BOQ import (`ImportToBOQAsync`) and PR generation (`ConvertBOQToPRAsync`) both require a project ID.
- Project search results use the `OdooProject` record with fields: Id, Name, Code, State, DateStart, DateEnd.
