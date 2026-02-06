# TASKS: US-005-01 — Convert BOQ to PR

> **Parent US**: [US-005-01](US-005-01-convert-boq-to-pr.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 9 | **Effort**: 5S + 4M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-04 (BOQ pushed to Odoo -- BOQ entries must exist before conversion)
- [ ] FR-003 Odoo Connection must be active (`IOdooService.IsConnected`)
- [ ] FR-002 AutoCAD Connection must be active if using header_id-based conversion path

## Acceptance Criteria
- [ ] AC-01: A "Convert BOQ to PR" button is visible on the Purchase Requisition page
- [ ] AC-02: Clicking the Convert button calls `IOdooService.ConvertBOQToPRAsync(projectId, boqEntryIds)` with the current project context
- [ ] AC-03: When no specific BOQ entry IDs are provided, all BOQ entries for the project are included in the conversion
- [ ] AC-04: User can select specific BOQ entries to convert (partial conversion)
- [ ] AC-05: A progress indicator is displayed while the Odoo backend processes the conversion request
- [ ] AC-06: On successful conversion, the PR list refreshes and the newly created PR is highlighted
- [ ] AC-07: On failure, the error message returned by the Odoo API is displayed to the user
- [ ] AC-08: The Convert button is disabled while a conversion is already in progress (IsConverting == true)
- [ ] AC-09: The Convert button is disabled when no BOQ data exists for the project, with a warning message shown

---

## TASK-005-01-01: Add "Convert BOQ to PR" button to page layout

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-01-02, TASK-005-01-06 |

### What to do
- Add a `Button` labeled "Convert BOQ to PR" in the Conversion Controls section of the page
- Style with brown color scheme (#5D4037 foreground, #3E2723 hover) to match Python sidebar button
- Bind `Command` to `{Binding ConvertBOQToPRCommand}`
- Bind `IsEnabled` to `{Binding IsConverting, Converter={StaticResource InverseBoolConverter}}` to prevent double-click (VR-005-007)
- Add a secondary disabled state when no BOQ data exists, bound to a `HasBOQData` property
- Position the button in the Conversion Controls area above the filter/sort bar per the wireframe

### How to verify
- [ ] Button is visible on the PurchaseRequisitionPage (AC-01)
- [ ] Button is disabled while `IsConverting == true` (AC-08)
- [ ] Button is disabled when no BOQ data exists with warning message (AC-09)

---

## TASK-005-01-02: Implement ConvertBOQToPRCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-005-01-01, TASK-005-01-03 |
| Blocks | TASK-005-01-06, TASK-005-01-07, TASK-005-01-08 |

### What to do
- Declare `IAsyncRelayCommand ConvertBOQToPRCommand` using `CommunityToolkit.Mvvm`
- Implement `ConvertBOQToPRAsync()` method that sets `IsConverting = true`, calls `_odooService.ConvertBOQToPRAsync(ProjectId.Value, selectedBoqEntryIds)`, and sets `IsConverting = false` in `finally`
- When `boqEntryIds` is null (no selection), pass null to include all BOQ entries (FR-005-003)
- When user has selected specific BOQ entries, pass their IDs for partial conversion (FR-005-004)
- Wire `CanExecute` to check `!IsConverting && ProjectId.HasValue && HasBOQData`
- Inject `IOdooService` via constructor

### How to verify
- [ ] Command calls `IOdooService.ConvertBOQToPRAsync(projectId, boqEntryIds)` with correct project context (AC-02)
- [ ] All BOQ entries included when no specific selection (AC-03)
- [ ] Partial conversion works when specific BOQ entries are selected (AC-04)

---

## TASK-005-01-03: Implement ConvertBOQToPRAsync in Odoo service

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Core/Odoo/IOdooService.cs`, `OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-01-02 |

### What to do
- The interface method `ConvertBOQToPRAsync(int projectId, IEnumerable<int>? boqEntryIds = null)` already exists in `IOdooService.cs`
- Verify the `OdooService.ConvertBOQToPRAsync()` implementation properly maps to Odoo's `boq2pr_v2` endpoint via `CallMethodAsync("boq.line", "convert_to_pr", methodParams)`
- Ensure the method passes `project_id` and optional `boq_entry_ids` in the request body
- Parse the response to extract the created PR ID, then fetch full PR data via `GetPurchaseRequisitionsAsync(projectId)`
- Throw meaningful exceptions on error responses (map Odoo `error_code`/`error_message` to C# exceptions)
- Handle null/empty result when no PR is created (return null)

### How to verify
- [ ] API call executes correctly against Odoo `boq2pr_v2` endpoint (AC-02)
- [ ] Returns `PREntry` on success, null when no PR created (AC-06, AC-07)
- [ ] Error responses from Odoo are propagated as exceptions with descriptive messages (AC-07)

---

## TASK-005-01-04: Add optional header_id extraction path via IAutoCADService

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs`, `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-01-05 |

### What to do
- Add `IReadOnlyList<int> GetLayoutsHeaderIds()` method to `IAutoCADService` interface (equivalent to Python's `get_layouts_header_id_to_pr()`)
- Implement in `AutoCADService`: iterate all document layouts (excluding "Model"), read block attributes, extract table data, collect `header_id` from valid tables (9 columns, "HEADER_ID" in column 7)
- Only include `header_id` when the layout has non-empty detail rows
- Return a flat list of header IDs: `List<int>`
- This is the alternative conversion path matching the Python pipeline; the project-based path is preferred for initial implementation

### How to verify
- [ ] Method extracts header_ids from AutoCAD layouts correctly (AC-02)
- [ ] Only non-Model layouts with valid tables are processed (AC-03)

---

## TASK-005-01-05: Wrap AutoCAD COM calls through IGUIProxy for thread safety

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-01-04 |
| Blocks | None |

### What to do
- When using the header_id-based conversion path, wrap `IAutoCADService.GetLayoutsHeaderIds()` calls through `IGUIProxy.ExecuteInGuiAsync("get_layouts_header_ids")`
- Register the `"get_layouts_header_ids"` action handler in the ViewModel or at application startup
- The handler calls `_autoCADService.GetLayoutsHeaderIds()` on the GUI/STA thread
- Extract the result from `GUIProxyResponse.Result` and cast to `IReadOnlyList<int>`
- Handle `GUIProxyResponse.Status == ProxyRequestStatus.TimedOut` and `ProxyRequestStatus.Failed`

### How to verify
- [ ] AutoCAD COM calls execute on GUI thread without COM threading exceptions (AC-02)
- [ ] Timeout and failure scenarios are handled gracefully (AC-07)

---

## TASK-005-01-06: Add progress indicator and conversion state management

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-01-01, TASK-005-01-02 |
| Blocks | None |

### What to do
- Declare `[ObservableProperty] private bool _isConverting;` with property change notification
- Declare `[ObservableProperty] private string _statusMessage;` for status text binding
- Set `IsConverting = true` and `StatusMessage = "Converting BOQ to Purchase Requisition..."` at start of `ConvertBOQToPRAsync()`
- Set `IsConverting = false` in `finally` block
- Notify `ConvertBOQToPRCommand` of `CanExecute` change when `IsConverting` changes
- Bind progress indicator in XAML to `{Binding IsConverting}` for visibility

### How to verify
- [ ] Progress indicator is visible during conversion (AC-05)
- [ ] Convert button is disabled during conversion (AC-08)

---

## TASK-005-01-07: Implement post-conversion PR list refresh and highlighting

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-01-02 |
| Blocks | None |

### What to do
- After successful `ConvertBOQToPRAsync`, call `await RefreshListAsync()` to reload the PR list from Odoo
- Set `SelectedPR = PurchaseRequisitions.FirstOrDefault(p => p.Id == result.Id)` to highlight the newly created PR
- Update `LastConversionTime = DateTime.Now` for display in the project context panel
- Ensure the ListView scrolls to and selects the newly created PR via `SelectedPR` binding

### How to verify
- [ ] PR list refreshes after successful conversion (AC-06)
- [ ] Newly created PR is highlighted/selected in the list (AC-06)

---

## TASK-005-01-08: Implement error handling for conversion failures

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-01-02 |
| Blocks | None |

### What to do
- Wrap `ConvertBOQToPRAsync` body in try/catch
- Catch `HttpRequestException` for network failures and set `StatusMessage` with timeout-specific message
- Catch `TaskCanceledException` for timeout and display: "The conversion request timed out. The Odoo server may be processing a large request. Please try again."
- Catch general `Exception` and set `StatusMessage = $"Conversion failed: {ex.Message}"` using the Odoo error_message
- When result is null (no PR created), set `StatusMessage = "Conversion completed but no Purchase Requisitions were created. The BOQ may already be fully converted."`
- Ensure `IsConverting = false` in `finally` block so the Convert button re-enables for retry

### How to verify
- [ ] Odoo API error message is displayed on failure (AC-07)
- [ ] Convert button re-enables after failure for retry (AC-08)

---

## TASK-005-01-09: Add BOQ entry selection UI for partial conversion

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | M |
| Depends On | TASK-005-01-01 |
| Blocks | None |

### What to do
- Add a collapsible section or popup showing available BOQ entries for the current project
- Fetch BOQ entries via `IOdooService.GetBOQEntriesAsync(projectId)` and bind to a `ListView` with checkboxes
- Each entry shows: Product Name, Quantity, Unit of Measure
- Add a "Select All / Deselect All" toggle
- Collect selected entry IDs into `SelectedBOQEntryIds` collection on the ViewModel
- Pass `SelectedBOQEntryIds` to `ConvertBOQToPRAsync` when the user has made a selection; pass null when "all" is selected

### How to verify
- [ ] User can select specific BOQ entries for partial conversion (AC-04)
- [ ] When no entries are explicitly selected, all entries are included (AC-03)

---

## Dependency Graph

```
TASK-005-01-03 (IOdooService PR method)
       |
       v
TASK-005-01-02 (ConvertBOQToPRCommand) <-- TASK-005-01-01 (XAML button)
       |
       +--------> TASK-005-01-06 (progress indicator)
       +--------> TASK-005-01-07 (post-conversion refresh)
       +--------> TASK-005-01-08 (error handling)

TASK-005-01-04 (GetLayoutsHeaderIds)
       |
       v
TASK-005-01-05 (IGUIProxy thread safety)

TASK-005-01-01 (XAML button) --> TASK-005-01-09 (BOQ selection UI)
```
