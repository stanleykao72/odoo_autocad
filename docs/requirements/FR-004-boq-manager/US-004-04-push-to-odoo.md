# US-004-04: Push to Odoo

## User Story
**As a** CAD Engineer,
**I want to** push validated BOQ entries to Odoo,
**So that** the bill of quantities is recorded in the ERP system.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Clicking "Push to Odoo" sends the validated layout data to Odoo via `IOdooService.ImportToBOQAsync()` (equivalent to `import2boq_v2` API)
- [ ] AC-02: The Push button is disabled until validation passes with zero Error-severity issues
- [ ] AC-03: The Push button is disabled while a push is already in progress (`IsPushing` is true)
- [ ] AC-04: The Push button is disabled when there are no extracted items (`TotalItems == 0`)
- [ ] AC-05: A progress indicator shows records processed out of total during the push operation
- [ ] AC-06: Upon push completion, a result summary displays: records created, records updated, records failed, and any error messages from the Odoo response
- [ ] AC-07: If Odoo is not connected, the Push button is disabled with message: "Odoo is not connected. Please connect to Odoo before pushing BOQ data."
- [ ] AC-08: If Odoo returns an `error_code`, the error message is displayed and ID writeback does not proceed
- [ ] AC-09: Partial push failures show per-record errors and allow retry for failed records
- [ ] AC-10: Network timeout during push shows partial completion status with count of processed items

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-018 | Provide "Push to Odoo" button that sends validated data via IOdooService.ImportToBOQAsync() | Must |
| FR-004-019 | Push disabled until validation passes with zero Error-severity issues | Must |
| FR-004-020 | After successful push, write returned header_id and detail_id back to AutoCAD tables | Must |
| FR-004-021 | Display progress indicator showing records processed out of total | Should |
| FR-004-022 | Display push result summary: records created, updated, failed, with error messages | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-04-01 | Add "Push to Odoo" button with CanExecute binding to validation state and connection state | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-04-02 | Implement PushToOdooCommand with pre-validation check and IBOQProcessor.PushToOdooAsync() call | `ViewModels/BOQManagerViewModel.cs` | L |
| TASK-004-04-03 | Implement ImportToBOQAsync in IOdooService (equivalent to import2boq_v2 API call) | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | L |
| TASK-004-04-04 | Parse Odoo response to extract header_id and detail per layout for writeback | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-04-05 | Add push progress reporting with IProgress<T> pattern | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-04-06 | Add push result summary panel XAML with status text and last push info | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-04-07 | Handle Odoo error_code responses and display user-facing error messages | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-04-08 | Handle partial failure and network timeout scenarios with retry capability | `ViewModels/BOQManagerViewModel.cs` | M |

## Dependencies
- Depends on: US-003-01 (Odoo connected), US-004-03 (validation passed)
- Blocks: US-004-05 (ID writeback depends on successful push)

## Notes
- The Push operation is the second step of the Python three-step workflow (`odoo_util.import2boq(layout_dict)`). In C#, it is triggered as a separate user action after explicit validation.
- The `CanExecute` for `PushToOdooCommand` is bound to: `!HasValidationErrors && !IsPushing && TotalItems > 0`.
- The Odoo API endpoint `job_working_plan_boq.import2boq_v2` returns a list with `header_id` and `detail` items per layout. This response is parsed to build the writeback list used by US-004-05.
- `PushStatusText` and `LastPushResult` observable properties provide real-time feedback to the user.
- `LastPushTime` records when the last push completed for audit trail display.
