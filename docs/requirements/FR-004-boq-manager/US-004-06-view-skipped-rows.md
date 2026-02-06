# US-004-06: View Skipped Rows

## User Story
**As a** CAD Engineer,
**I want to** see which rows were skipped during extraction and why,
**So that** I can fix issues in the drawing if needed.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Rows skipped because both `qty` and `product_no` are empty/whitespace are counted and reported in the summary panel under "Skipped Items"
- [ ] AC-02: Tables skipped because they fail the legal table check (not 9 columns or missing HEADER_ID marker) are counted separately
- [ ] AC-03: The summary panel displays the total Skipped Items count with a gray background
- [ ] AC-04: A tooltip or expandable section on the Skipped count shows a breakdown: number of empty rows skipped, number of illegal tables skipped
- [ ] AC-05: The user can identify which layouts contained skipped rows or illegal tables

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-004 | Skip rows where both qty and product_no are empty/whitespace after MText unformatting | Must |
| FR-004-010 | Display summary panel showing Skipped Items count among other metrics | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-06-01 | Track skipped row count during extraction (empty qty + product_no rows) | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-06-02 | Track skipped table count during extraction (illegal tables) | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-06-03 | Display SkippedItems in summary panel with gray background | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-06-04 | Add tooltip on Skipped count showing breakdown (empty rows vs illegal tables) | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-06-05 | Store per-layout skip information for user reference | `ViewModels/BOQManagerViewModel.cs` | M |

## Dependencies
- Depends on: US-004-01 (extraction process generates skip information)
- Blocks: None

## Notes
- Two categories of skipping occur during extraction:
  1. **Row-level skips**: Rows where both `qty` (column 6) and `product_no` (column 1) are empty/whitespace after MText unformatting (VR-004-005).
  2. **Table-level skips**: Tables that fail the legal table check -- either they do not have exactly 9 columns (VR-004-001) or they lack the `HEADER_ID` marker at cell (0, 7) (VR-004-002).
- Layouts named "Model" are excluded entirely (VR-004-004) and are not counted as skipped.
- The Python reference silently skips these items. The C# implementation adds visibility into what was skipped and why, which is a usability improvement over the Python version.
