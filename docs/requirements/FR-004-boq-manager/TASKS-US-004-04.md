# TASKS: US-004-04 — Push to Odoo

> **Parent US**: [US-004-04](US-004-04-push-to-odoo.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 3S + 3M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (Odoo connected — `IOdooService.IsConnected` must return true)
- [ ] US-004-03 (Validation passed with zero Error-severity issues)

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

---

## TASK-004-04-01: Add "Push to Odoo" button with CanExecute binding to validation state and connection state

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Push to Odoo" `Button` in the push controls section of `BOQManagerPage.xaml`
- Bind `Command="{Binding PushToOdooCommand}"`
- The `CanExecute` is managed by the ViewModel command; no additional `IsEnabled` binding needed
- Add a `TextBlock` for Odoo connection warning message when `!IsOdooConnected`: "Odoo is not connected. Please connect to Odoo before pushing BOQ data."
- Add push status text `TextBlock` bound to `{Binding PushStatusText}` showing readiness state
- Add last push info `TextBlock` bound to `{Binding LastPushResult}` showing result of previous push
- Position the button in the bottom push controls bar, after "Validate All"

### How to verify
- [ ] "Push to Odoo" button renders and is bound to PushToOdooCommand (AC-01)
- [ ] Button is disabled when Odoo is disconnected, showing warning message (AC-07)
- [ ] Push status text and last push result are displayed (AC-06)

---

## TASK-004-04-02: Implement PushToOdooCommand with pre-validation check and IBOQProcessor.PushToOdooAsync() call

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | L |
| Depends On | TASK-004-04-01 |
| Blocks | TASK-004-04-04, TASK-004-04-05, TASK-004-04-07, TASK-004-04-08 |

### What to do
- Define `PushToOdooCommand` as `IAsyncRelayCommand` using `AsyncRelayCommand(ExecutePushAsync, CanExecutePush)`
- `CanExecutePush()` returns `!HasValidationErrors && !IsPushing && TotalItems > 0 && IsOdooConnected`
- In `ExecutePushAsync()`: set `IsPushing = true`, `PushStatusText = "Pushing to Odoo..."`
- Convert `AllEntries` to `IEnumerable<BOQEntry>` with proper field mapping (ProductNo -> ProductId via mapping, Qty -> Quantity, etc.)
- Set `ProjectId` on all entries from the resolved project context
- Call `_boqProcessor.PushToOdooAsync(entries)` which internally validates then calls `IOdooService.ImportToBOQAsync()`
- Store the returned `SyncResult` for response parsing and writeback
- On success: set `LastPushResult` with summary text and `LastPushTime = DateTime.Now`
- On failure: display error messages from `SyncResult.Errors` in `PushStatusText`
- Set `IsPushing = false` in a `finally` block

### How to verify
- [ ] PushToOdooCommand calls `IBOQProcessor.PushToOdooAsync()` with validated entries (AC-01)
- [ ] Push is disabled until validation passes with zero errors (AC-02)
- [ ] Push is disabled while already pushing (AC-03)
- [ ] Push is disabled when TotalItems is 0 (AC-04)

---

## TASK-004-04-03: Implement ImportToBOQAsync in IOdooService (equivalent to import2boq_v2 API call)

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-004-04-02 |

### What to do
- Implement `ImportToBOQAsync(IEnumerable<BOQEntry> entries)` in `OdooService` class
- Build the API request payload matching the Python `import2boq_v2` format: a layout dict structure with `'all'` key containing layout entries with header-level metadata and `detail` arrays
- Call the Odoo REST endpoint `job_working_plan_boq.import2boq_v2` via the HTTP client
- Parse the response which returns a list with `header_id` and `detail` (containing `detail_id` per row matched by `product_no`)
- Handle error responses: check for `error_code` field in the response JSON
- Build and return `SyncResult` with `RecordsCreated`, `RecordsUpdated`, `RecordsFailed` counts from the response
- Implement request timeout handling with `HttpClient.Timeout` or `CancellationToken`
- Add structured logging for the request/response cycle

### How to verify
- [ ] API call matches the `import2boq_v2` endpoint format (AC-01)
- [ ] Response is parsed to extract header_id and detail_id values for writeback (AC-01)
- [ ] Error responses with error_code are properly detected (AC-08)

---

## TASK-004-04-04: Parse Odoo response to extract header_id and detail per layout for writeback

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-04-02 |
| Blocks | None |

### What to do
- After `PushToOdooAsync()` returns successfully, parse the `SyncResult` to extract the writeback data
- The Odoo response contains per-layout data: `header_id` and a `detail` list with `detail_id` and `product_no` for each row
- Build a writeback data structure (e.g., `List<WritebackLayout>`) containing:
  - `LayoutName`, `HeaderId`, and a list of `(ProductNo, DetailId)` tuples
- Store this writeback data for use by US-004-05 (ID Writeback)
- If the response does not contain writeback data (error case), log a warning and skip writeback preparation
- Update `BOQEntryRow.DetailId` and `BOQLayoutGroup.HeaderId` with the returned values for immediate UI feedback

### How to verify
- [ ] header_id and detail_id values are extracted from the Odoo response per layout (AC-01)
- [ ] Writeback data structure is ready for US-004-05 consumption (AC-06)

---

## TASK-004-04-05: Add push progress reporting with IProgress<T> pattern

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-04-02 |
| Blocks | None |

### What to do
- Create a `Progress<(int current, int total, string message)>` instance in `ExecutePushAsync()`
- In the progress callback, update `PushStatusText` (e.g., `"Processing 15/42..."`) and a push progress percent property
- If the push API is batch-based (per layout), report progress after each layout is processed
- Update `ExtractionProgressPercent` or a separate `PushProgressPercent` property for the progress bar binding
- Ensure progress updates are dispatched to the UI thread

### How to verify
- [ ] During push, progress shows records processed out of total (AC-05)
- [ ] Progress text updates as records are processed

---

## TASK-004-04-06: Add push result summary panel XAML with status text and last push info

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a push result section in the push controls area of `BOQManagerPage.xaml`
- Include a `TextBlock` bound to `{Binding PushStatusText}` for current push state (e.g., "Ready (42 valid entries, 2 errors to resolve)")
- Include a `TextBlock` bound to `{Binding LastPushResult}` for last push outcome (e.g., "Last Push: 2026-02-05 16:30 - 40 records created, 0 failed")
- Include a push `ProgressBar` bound to push progress percent, visible only when `IsPushing` is true
- Style the result summary with appropriate colors: green for success, red for failure, neutral for pending

### How to verify
- [ ] Push result summary displays records created, updated, failed (AC-06)
- [ ] Status text shows readiness state and last push result (AC-06)

---

## TASK-004-04-07: Handle Odoo error_code responses and display user-facing error messages

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-04-02 |
| Blocks | None |

### What to do
- In `ExecutePushAsync()`, after receiving `SyncResult`, check if `Success` is false
- If `SyncResult.Errors` contains entries, format them as user-facing messages:
  - `PushStatusText = $"Odoo returned an error: {string.Join("; ", syncResult.Errors)}"`
- If Odoo returns an `error_code` (detected in `OdooService.ImportToBOQAsync()`), surface it in the error list
- When an error occurs, do NOT proceed with ID writeback (set a flag `shouldWriteBack = false`)
- Display the error in the push result summary area with red styling
- Log the error details via `_logger.LogError()`

### How to verify
- [ ] Odoo error_code responses display the error message to the user (AC-08)
- [ ] ID writeback does not proceed when push returns errors (AC-08)

---

## TASK-004-04-08: Handle partial failure and network timeout scenarios with retry capability

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | TASK-004-04-02 |
| Blocks | None |

### What to do
- Wrap the `PushToOdooAsync()` call in a try-catch for `TaskCanceledException` (timeout) and `HttpRequestException` (network errors)
- On timeout: set `PushStatusText = $"Connection to Odoo timed out during BOQ push. {processedItems}/{totalItems} items were processed."`
- Track the number of items processed before the timeout for partial completion reporting
- On partial failure (some records succeed, some fail): parse `SyncResult` to identify failed records
- Add a "Retry Failed" button or command that retries only the failed records
- Store failed entry indices in a `List<int> _failedEntryIndices` for retry targeting
- Display per-record errors from `SyncResult.Errors` in the push result panel

### How to verify
- [ ] Partial push failures show per-record errors (AC-09)
- [ ] Network timeout shows partial completion with processed count (AC-10)
- [ ] Retry capability is available for failed records (AC-09)

---

## Dependency Graph
```
TASK-004-04-03 (ImportToBOQAsync Service)
       │
       └──▶ TASK-004-04-02 (PushToOdooCommand ViewModel)
                   │
                   ├──▶ TASK-004-04-04 (Parse Response for Writeback)
                   ├──▶ TASK-004-04-05 (Push Progress Reporting)
                   ├──▶ TASK-004-04-07 (Error Code Handling)
                   └──▶ TASK-004-04-08 (Partial Failure / Timeout)

TASK-004-04-01 (Push Button XAML) ─── independent
TASK-004-04-06 (Push Result Summary XAML) ─── independent

Tasks 01, 03, 06 can begin in parallel.
Task 02 depends on 03 (service implementation needed).
Tasks 04, 05, 07, 08 all depend on 02 (command must exist).
```
