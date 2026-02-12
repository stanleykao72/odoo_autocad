# US-009-01: Fetch Odoo Parameters

## User Story
**As a** CAD Engineer,
**I want to** fetch material/setup/color options from Odoo via API,
**So that** I have up-to-date dropdown data for my drawing parameters.

## Parent Feature
- **FR**: [FR-009-parameter-config](FR-009-parameter-config.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: `GetSetupViaApiAsync(setupName)` fetches setup values from Odoo via Swagger `get_setup_v2` endpoint
- [ ] AC-02: `GetColorsViaApiAsync(projectId)` fetches color options from Odoo via Swagger `get_color_v2` endpoint
- [ ] AC-03: Setup values are returned as `IReadOnlyList<OdooSetupValue>` with `Value` and `SetupName` properties
- [ ] AC-04: Color values are returned as `IReadOnlyList<OdooColor>` with `Name`, `ColorNo`, and `ProjectId` properties
- [ ] AC-05: Both methods use BasicAuth + PATCH pattern consistent with existing Swagger methods
- [ ] AC-06: API failures return empty lists and log error details (no exceptions thrown to caller)
- [ ] AC-07: `GetProductsViaApiAsync()` (existing) is reused for material dropdown data

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-009-001 | Page SHALL fetch product options from Odoo via `get_product_v2` Swagger endpoint | Must |
| FR-009-002 | Page SHALL fetch setup values via `get_setup_v2` Swagger endpoint | Must |
| FR-009-003 | Page SHALL fetch project-specific color options via `get_color_v2` Swagger endpoint | Must |
| FR-009-004 | All API calls SHALL use BasicAuth + Swagger PATCH pattern | Must |
| FR-009-006 | API failures SHALL show user-friendly error messages without crashing | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-009-01-01 | Define OdooSetupValue and OdooColor record types | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |
| TASK-009-01-02 | Add GetSetupViaApiAsync and GetColorsViaApiAsync to IOdooService interface | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |
| TASK-009-01-03 | Implement GetSetupViaApiAsync in OdooService | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-009-01-04 | Implement GetColorsViaApiAsync in OdooService | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-009-01-05 | Unit tests for GetSetupViaApiAsync | `tests/.../OdooServiceTests.cs` | M |
| TASK-009-01-06 | Unit tests for GetColorsViaApiAsync | `tests/.../OdooServiceTests.cs` | S |

## Dependencies
- Depends on: Sprint 3 infrastructure (IOdooService, Swagger API pattern, BasicAuth)
- Depends on: `GetProductsViaApiAsync()` (existing, Sprint 5)
- Blocks: US-009-02 (form UI needs dropdown data)

## Notes
- The Python `get_product_v2` uses filter `categ_id child_of 27, active=True`. The C# `GetProductsViaApiAsync()` already implements this.
- The Python `get_setup_v2(setup_name)` uses filter `setup_name = {name}`. The C# implementation passes `setup_name` as a parameter in the Swagger body.
- The Python `get_color_v2(project_id)` uses filter `job_project_id = {project_id}`. The C# implementation passes `project_id` as a parameter.
- All three methods use the `callMethodForJobWorkingPlanBoqModel` Swagger endpoint with `user_token` in kwargs.
- The existing `GetProductsViaApiAsync` pattern (PATCH + BasicAuth + `ResolveSwaggerEndpoint`) should be followed exactly for the two new methods.
