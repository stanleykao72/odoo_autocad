# US-004-03: Validate Against Products

## User Story
**As a** CAD Engineer,
**I want to** validate BOQ entries against Odoo product catalog,
**So that** I know which products are recognized and which need mapping.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Clicking "Validate" checks all extracted BOQ entries against Odoo product catalog and business rules
- [ ] AC-02: Each entry's `product_no` is verified to map to a valid Odoo product ID (case-insensitive matching via local mapping cache first, then Odoo query)
- [ ] AC-03: Each entry's `qty` is verified to be a positive numeric value; non-numeric or negative values produce Error severity; zero produces Warning severity
- [ ] AC-04: A valid project context (non-zero `ProjectId` from `pr_no` lookup) is verified before push is allowed
- [ ] AC-05: Per-entry validation errors include entry index, field name, error message, and severity level (Error or Warning)
- [ ] AC-06: Rows with non-empty `product_no` but empty `qty` produce a Warning-severity validation error
- [ ] AC-07: Unresolved `product_no` values produce Error severity (or Warning if `AutoCreateProducts` is enabled)
- [ ] AC-08: Validation results update the row-level ValidationStatus and ValidationMessage properties on each BOQEntryRow
- [ ] AC-09: Summary counts (ValidItems, WarningItems, ErrorItems) are recalculated after validation

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-012 | Provide "Validate" button that checks all entries against Odoo products and business rules | Must |
| FR-004-013 | Verify each entry has non-empty product_no mappable to a valid Odoo product ID | Must |
| FR-004-014 | Verify each entry has a positive numeric qty value | Must |
| FR-004-015 | Verify valid project context exists (non-zero ProjectId from pr_no lookup) | Must |
| FR-004-016 | Produce per-entry error details including index, field, message, and severity | Must |
| FR-004-017 | Display validation results in panel below DataGrid with clickable errors that highlight corresponding row | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-03-01 | Add "Validate All" button to BOQ Manager page bound to ValidateCommand | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-03-02 | Implement ValidateCommand that calls IBOQProcessor.ValidateBOQAsync(entries) | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-03-03 | Implement ValidateBOQAsync with product_no resolution (local cache then Odoo query) | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | L |
| TASK-004-03-04 | Implement qty validation rules: positive numeric required, zero=Warning, negative/non-numeric=Error | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | M |
| TASK-004-03-05 | Implement project context validation via IOdooService.GetProjectAsync() | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |
| TASK-004-03-06 | Build BOQValidationResult/BOQValidationError model population with per-entry details | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | M |
| TASK-004-03-07 | Update BOQEntryRow.ValidationStatus and ValidationMessage after validation completes | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-03-08 | Recalculate summary counts (ValidItems, WarningItems, ErrorItems) post-validation | `ViewModels/BOQManagerViewModel.cs` | S |

## Dependencies
- Depends on: US-004-01 (extraction must provide data to validate)
- Blocks: US-004-04 (push requires validation with zero errors)

## Notes
- Validation is a new step added in the C# implementation that does not exist in the Python reference. Python pushes directly without explicit validation.
- Product resolution follows a two-tier approach: first check `IBOQProcessor.GetProductMappings()` (case-insensitive via `StringComparer.OrdinalIgnoreCase`), then query Odoo via `IOdooService.SearchProductsAsync()`.
- Validation rules are defined in FR-004 Section 8: VR-004-005 through VR-004-013.
- The `BOQGenerationOptions.ValidateProducts` flag controls whether unresolved products are Error (true) or Warning (false with `AutoCreateProducts`).
