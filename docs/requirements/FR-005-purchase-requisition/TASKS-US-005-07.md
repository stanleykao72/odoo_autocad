# TASKS: US-005-07 — Conversion Feedback

> **Parent US**: [US-005-07](US-005-07-conversion-feedback.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 7S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-005-01 (Conversion must be triggered before feedback can be displayed)
- [ ] `StatusMessage` property bound in the ViewModel
- [ ] `ConvertBOQToPRAsync()` method implemented in the ViewModel

## Acceptance Criteria
- [ ] AC-01: On successful conversion, the status area displays a success message including the newly created PR reference (e.g., "Successfully created PR: PR-2025-006")
- [ ] AC-02: On successful conversion, the PR list refreshes and the newly created PR is visually highlighted/selected
- [ ] AC-03: On failure, the status area displays the error message returned by the Odoo API (e.g., "Conversion failed: {error_message}")
- [ ] AC-04: When conversion completes but no PR was created, an informational message is shown: "Conversion completed but no Purchase Requisitions were created. The BOQ may already be fully converted."
- [ ] AC-05: On network timeout, a specific message is shown: "The conversion request timed out. The Odoo server may be processing a large request. Please try again."
- [ ] AC-06: The `LastConversionTime` is updated on successful conversion and displayed in the project context panel
- [ ] AC-07: Error states provide a retry option so the user can attempt conversion again without additional steps

---

## TASK-005-07-01: Add status message area in conversion controls section

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-07-07 |

### What to do
- Add a `TextBlock` or `Border`+`TextBlock` in the Conversion Controls section below the Convert and Refresh buttons
- Bind `Text` to `{Binding StatusMessage}`
- The status area should span the full width of the conversion controls section
- Include a left-side color accent bar or icon area that changes based on status type (success/error/info)
- Bind `Visibility` to show only when `StatusMessage` is not empty
- Set `TextWrapping = Wrap` to handle longer error messages

### How to verify
- [ ] Status message area is visible in the conversion controls section (AC-01, AC-03)
- [ ] Message text updates based on conversion result (AC-01, AC-03, AC-04)

---

## TASK-005-07-02: Implement success feedback logic

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `ConvertBOQToPRAsync()`, after successful conversion where `result != null`:
  - Set `StatusMessage = $"Successfully created PR: {result.Reference}"`
  - Set `StatusMessageType = "success"` (for styling in TASK-005-07-07)
  - Update `LastConversionTime = DateTime.Now`
  - Call `await RefreshListAsync()` to reload the PR list
  - Set `SelectedPR = PurchaseRequisitions.FirstOrDefault(p => p.Id == result.Id)` to highlight the new PR
- The newly created PR should scroll into view and be visually selected in the list

### How to verify
- [ ] Success message includes the new PR reference (AC-01)
- [ ] PR list refreshes and new PR is highlighted (AC-02)
- [ ] LastConversionTime is updated (AC-06)

---

## TASK-005-07-03: Implement failure feedback logic

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `ConvertBOQToPRAsync()`, in the `catch` block for general exceptions:
  - Set `StatusMessage = $"Conversion failed: {ex.Message}"` using the Odoo error_message
  - Set `StatusMessageType = "error"` (for red styling)
- Map common Odoo error responses to user-friendly messages per FR-005 section 9:
  - No BOQ data: "No BOQ entries found for this project. Push parameters to BOQ first before converting to PR."
  - No header_ids: "No valid layout tables found in the current drawing."
  - API error: "Conversion failed: {error_message from Odoo}"
- Ensure `IsConverting = false` in `finally` to re-enable the Convert button for retry

### How to verify
- [ ] Odoo API error message displayed in status area (AC-03)
- [ ] Convert button re-enabled for retry after failure (AC-07)

---

## TASK-005-07-04: Implement empty result feedback

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `ConvertBOQToPRAsync()`, when result is null (conversion returned no PR):
  - Set `StatusMessage = "Conversion completed but no Purchase Requisitions were created. The BOQ may already be fully converted."`
  - Set `StatusMessageType = "info"` (for blue styling)
- This scenario occurs when the Odoo backend determines that all BOQ entries have already been converted
- Do not show an error tone -- this is informational, not a failure

### How to verify
- [ ] Informational message shown when no PR is created (AC-04)
- [ ] Message tone is informational, not error (AC-04)

---

## TASK-005-07-05: Implement network timeout detection

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `ConvertBOQToPRAsync()`, catch `TaskCanceledException` and `OperationCanceledException` specifically
- Check if the exception is due to a timeout (not user cancellation)
- Set `StatusMessage = "The conversion request timed out. The Odoo server may be processing a large request. Please try again."`
- Set `StatusMessageType = "error"`
- Also catch `HttpRequestException` for general network failures and set appropriate message
- Ensure the Convert button re-enables for retry

### How to verify
- [ ] Timeout-specific message shown on network timeout (AC-05)
- [ ] User can retry after timeout (AC-07)

---

## TASK-005-07-06: Add LastConversionTime display in project context panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Project Context panel at the top of the page, add a label: "Last Conversion:"
- Bind to `{Binding LastConversionTime, StringFormat='HH:mm'}` to show time in 24-hour format (e.g., "14:30")
- Declare `[ObservableProperty] private DateTime? _lastConversionTime;` in the ViewModel
- When `LastConversionTime` is null (no conversion yet), display "--" or "Never"
- Use a value converter or FallbackValue for null handling
- Position next to the PR count and project name as shown in the wireframe

### How to verify
- [ ] LastConversionTime displayed in the project context panel (AC-06)
- [ ] Time updates after each successful conversion (AC-06)

---

## TASK-005-07-07: Style status message area with visual differentiation

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-07-01 |
| Blocks | None |

### What to do
- Add a `StatusMessageType` string property to the ViewModel with values: "success", "error", "info", "none"
- Create `DataTrigger` or `IValueConverter` (`StatusTypeToColorConverter`) to map type to visual style:
  - "success": green text (#2E7D32), green left accent bar, light green background (#E8F5E9)
  - "error": red text (#C62828), red left accent bar, light red background (#FFEBEE)
  - "info": blue text (#1565C0), blue left accent bar, light blue background (#E3F2FD)
  - "none": hidden
- Apply the style to the status message `Border` and `TextBlock` created in TASK-005-07-01
- Include an icon (check mark for success, X for error, info circle for info) using a Path or Symbol

### How to verify
- [ ] Success messages display with green styling (AC-01)
- [ ] Error messages display with red styling (AC-03)
- [ ] Info messages display with blue styling (AC-04)

---

## Dependency Graph

```
TASK-005-07-01 (status area XAML) --> TASK-005-07-07 (styling)

TASK-005-07-02 (success feedback) ---------> independent
TASK-005-07-03 (failure feedback) ---------> independent
TASK-005-07-04 (empty result feedback) ----> independent
TASK-005-07-05 (timeout detection) --------> independent
TASK-005-07-06 (LastConversionTime) -------> independent
```
