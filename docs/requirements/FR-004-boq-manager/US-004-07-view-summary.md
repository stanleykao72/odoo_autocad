# US-004-07: View Summary

## User Story
**As a** Project Manager,
**I want to** view BOQ generation summary (total, processed, skipped),
**So that** I can gauge data quality at a glance.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: The summary panel displays Total Layouts processed
- [ ] AC-02: The summary panel displays Total Items extracted
- [ ] AC-03: The summary panel displays Valid Items count with green background
- [ ] AC-04: The summary panel displays Invalid Items count (Warnings with yellow background, Errors with red background)
- [ ] AC-05: The summary panel displays Skipped Items count with gray background
- [ ] AC-06: Summary counts are updated automatically after both extraction and validation complete
- [ ] AC-07: When no data has been extracted, the summary panel shows all zeros with an appropriate message

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-010 | Display summary panel showing: Total Layouts processed, Total Items extracted, Valid Items count, Invalid Items count, and Skipped Items count | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-07-01 | Implement summary bar XAML using UniformGrid with 5 cells and color-coded backgrounds | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-07-02 | Bind summary count properties (TotalItems, ValidItems, WarningItems, ErrorItems, SkippedItems) to UI | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-07-03 | Implement summary calculation method that tallies counts from AllEntries collection | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-07-04 | Trigger summary recalculation after extraction completes | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-07-05 | Trigger summary recalculation after validation completes | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-07-06 | Add LayoutsFound count property and bind to summary panel | `ViewModels/BOQManagerViewModel.cs` | S |

## Dependencies
- Depends on: US-004-01 (extraction provides total and skipped counts), US-004-03 (validation provides valid/warning/error counts)
- Blocks: None

## Notes
- The summary bar is a horizontal strip using `UniformGrid` with 5 cells. Each cell has a bold count label and a descriptive sub-label, with colored backgrounds: green for valid, yellow for warnings, red for errors, gray for skipped, and a neutral color for totals.
- Summary counts are derived from the ViewModel observable properties: `TotalItems`, `ValidItems`, `WarningItems`, `ErrorItems`, `SkippedItems`, and `LayoutsFound`.
- Before extraction, all counts should be 0. After extraction but before validation, `ValidItems` / `WarningItems` / `ErrorItems` should show as "Pending" or 0 until validation is run.
- The `LayoutsFound` property shows the number of non-Model layouts discovered (regardless of whether they contained legal tables).
