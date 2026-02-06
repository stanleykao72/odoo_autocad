# US-002-05: PR Project Info

## User Story
**As a** CAD Engineer,
**I want to** view extracted PR number and project info,
**So that** I can verify the correct project context.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: The page displays the extracted PR Number from the active layout in the Connection Status panel
- [ ] AC-02: The page displays the Project Name extracted from layout block attributes
- [ ] AC-03: The page displays the Job Working Plan Name extracted from layout block attributes
- [ ] AC-04: PR number and project info are shown in the format "PR No: PR-2025-001 | Project: Steel Frame Phase 2" as illustrated in the wireframe
- [ ] AC-05: If the PR number is not found in Odoo projects, a warning is displayed: "PR number '{pr_no}' not found in Odoo projects." but extraction continues
- [ ] AC-06: The ProjectId property is populated when a matching Odoo project is found via IOdooService.GetProjectAsync()
- [ ] AC-07: PR number and project info update when parameters are extracted from a different layout

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-013 | Page SHALL display extracted PR Number from the active layout | Must |
| FR-002-014 | Page SHALL display Project Name, Job Working Plan Name from layout attributes | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-05-01 | Add PR Number, Project Name, and Job Working Plan Name display fields to the Connection Status panel | `Views/Pages/AutoCADPage.xaml` | S |
| TASK-002-05-02 | Implement PRNumber, ProjectName, JobWorkingPlanName, and ProjectId properties in ViewModel | `ViewModels/AutoCADViewModel.cs` | S |
| TASK-002-05-03 | Implement process_pr_no equivalent in AutoCADService to extract PR number and project from layout | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-002-05-04 | Integrate IOdooService.GetProjectAsync() lookup to resolve PR number to Odoo ProjectId | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-002-05-05 | Add warning display when PR number is not found in Odoo projects | `ViewModels/AutoCADViewModel.cs` | S |

## Dependencies
- Depends on: US-002-03 (parameter extraction provides the PR number and project attributes)
- Blocks: None

## Notes
- The Python function `process_pr_no(layout)` extracts the PR number and project information from a layout's block attributes. The C# implementation should provide equivalent functionality.
- PR number is one of the 10 standard tag attributes (pr_no) extracted from layout blocks.
- The project context (ProjectId) is obtained by looking up the PR number via the Odoo service. This lookup is used later by the BOQ processing workflow.
- The wireframe shows PR and project info in the Connection Status panel: "PR No: PR-2025-001 | Project: Steel Frame Phase 2".
- Even if the Odoo lookup fails, the extracted PR and project name from AutoCAD should still be displayed. The Odoo lookup failure is a warning, not a blocking error.
