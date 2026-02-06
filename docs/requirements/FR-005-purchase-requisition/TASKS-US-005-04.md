# TASKS: US-005-04 — Submit PR

> **Parent US**: [US-005-04](US-005-04-submit-pr.md)
> **Parent FR**: [FR-005](FR-005-purchase-requisition.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 5S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-005-02 (PR list must be loaded to select a PR)
- [ ] US-005-03 (PR detail view must be visible to access Submit button)
- [ ] `IOdooService.SubmitPRAsync(int prId)` interface method defined
- [ ] `PRDisplayItem.CanSubmit` property computed correctly (State == "draft" && LineCount > 0)

## Acceptance Criteria
- [ ] AC-01: A "Submit for Approval" button is visible in the PR detail panel for PRs in "draft" state
- [ ] AC-02: The Submit button is disabled for PRs that are not in "draft" state
- [ ] AC-03: Clicking Submit triggers a confirmation dialog: "Submit PR {reference} for approval?"
- [ ] AC-04: On confirmation, `IOdooService.SubmitPRAsync(prId)` is called
- [ ] AC-05: On successful submission, the PR state updates to "submitted" in the UI immediately
- [ ] AC-06: On failure, an error message is displayed and the PR remains in "draft" state
- [ ] AC-07: The Submit button is disabled for PRs that have no line items, with a warning: "PR has no line items and cannot be submitted"
- [ ] AC-08: A loading indicator is shown on the Submit button while the submission is in progress

---

## TASK-005-04-01: Add "Submit for Approval" button to PR detail panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-005-04-02, TASK-005-04-07 |

### What to do
- Add a `Button` labeled "Submit for Approval" below the PR detail DataGrid
- Bind `Command` to `{Binding SubmitPRCommand}`
- Bind `IsEnabled` to `{Binding SelectedPR.CanSubmit}` -- this is `true` only when `State == "draft"` and `LineCount > 0`
- Bind `Visibility` to `{Binding SelectedPR, Converter={StaticResource NullToVisibilityConverter}}` so it only shows when a PR is selected
- Add a tooltip for disabled state: "PR must be in Draft state with line items to submit" when `CanSubmit == false`
- When `SelectedPR.LineCount == 0`, show warning text: "PR has no line items and cannot be submitted"

### How to verify
- [ ] Submit button is visible for selected PRs in the detail panel (AC-01)
- [ ] Button is disabled for non-draft PRs (AC-02)
- [ ] Button is disabled for PRs with no line items with warning message (AC-07)

---

## TASK-005-04-02: Implement SubmitPRCommand with CanExecute logic

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-005-04-01, TASK-005-04-03 |
| Blocks | TASK-005-04-04, TASK-005-04-05, TASK-005-04-06 |

### What to do
- Declare `IAsyncRelayCommand SubmitPRCommand` using `CommunityToolkit.Mvvm`
- Implement `SubmitPRAsync()` method that:
  1. Shows confirmation dialog (TASK-005-04-04)
  2. If confirmed, calls `_odooService.SubmitPRAsync(SelectedPR.Id)`
  3. On success, updates PR state (TASK-005-04-05)
  4. On failure, shows error (TASK-005-04-06)
- Set `CanExecute` to return `SelectedPR != null && SelectedPR.CanSubmit && !IsSubmitting`
- Declare `[ObservableProperty] private bool _isSubmitting;` for submission-in-progress state
- Notify `SubmitPRCommand` of `CanExecute` change when `SelectedPR` or `IsSubmitting` changes

### How to verify
- [ ] Command calls `IOdooService.SubmitPRAsync(prId)` on confirmation (AC-04)
- [ ] CanExecute is false when PR is not draft or has no lines (AC-02, AC-07)

---

## TASK-005-04-03: Implement SubmitPRAsync in Odoo service

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Core/Odoo/IOdooService.cs`, `OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-005-04-02 |

### What to do
- The interface method `SubmitPRAsync(int prId)` already exists in `IOdooService.cs`
- Verify/enhance `OdooService.SubmitPRAsync()` implementation that calls `CallMethodAsync("purchase.requisition", "action_submit", new { id = prId })`
- Return `true` on success, throw descriptive exception on failure
- Handle backend rejection: if the Odoo server returns an error (e.g., validation failure), extract the server-provided reason and throw with that message
- Handle the case where the PR has already been submitted (idempotency check)
- Log the submission attempt and result via `_logger`

### How to verify
- [ ] API call executes correctly against Odoo `action_submit` endpoint (AC-04)
- [ ] Returns true on success (AC-05)
- [ ] Backend rejection reasons are propagated as exception messages (AC-06)

---

## TASK-005-04-04: Implement confirmation dialog before submission

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-04-02 |
| Blocks | None |

### What to do
- Before calling `SubmitPRAsync`, show a confirmation dialog with message: "Submit PR {SelectedPR.Reference} for approval?"
- Use the application's standard dialog service (e.g., `MessageBox.Show` or a custom `IDialogService`)
- Dialog should have "Yes" and "No" buttons (or "Submit" and "Cancel")
- If user cancels, abort the submission without any state changes
- Include the PR reference in the dialog message for clarity

### How to verify
- [ ] Confirmation dialog appears before submission (AC-03)
- [ ] Canceling the dialog aborts submission (AC-03)

---

## TASK-005-04-05: Update PR state in UI after successful submission

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-04-02 |
| Blocks | None |

### What to do
- After `SubmitPRAsync` returns `true`, update the `PRDisplayItem` properties in-place:
  - Set `SelectedPR.State = "submitted"`
  - Set `SelectedPR.StateDisplay = "Submitted"`
  - Set `SelectedPR.StateBadgeBackground` = new SolidColorBrush from "#BBDEFB"
  - Set `SelectedPR.StateBadgeForeground` = new SolidColorBrush from "#1565C0"
  - Set `SelectedPR.CanSubmit = false`
- Fire property change notifications on the `PRDisplayItem` so the badge updates in both the list and detail views
- Ensure `SubmitPRCommand.NotifyCanExecuteChanged()` is called since `CanSubmit` is now false
- Optionally call `UpdateSummaryCounts()` to update the summary bar (US-005-08)

### How to verify
- [ ] PR state updates from "draft" to "submitted" in the UI (AC-05)
- [ ] State badge changes from gray to blue in both list and detail views (AC-05)
- [ ] Submit button becomes disabled after successful submission (AC-02)

---

## TASK-005-04-06: Implement error handling for submission failures

| Field | Value |
|-------|-------|
| Target | `ViewModels/PurchaseRequisitionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-005-04-02 |
| Blocks | None |

### What to do
- Wrap `SubmitPRAsync` call in try/catch
- Catch `HttpRequestException` for network failures: set `StatusMessage = "Failed to submit PR {reference} for approval: {error details}"`
- Catch general `Exception`: display error message from the Odoo backend
- On failure, the PR must remain in "draft" state -- do not change any UI state
- Set `IsSubmitting = false` in the `finally` block
- Display the error in the `StatusMessage` area with red styling

### How to verify
- [ ] Error message displayed on submission failure (AC-06)
- [ ] PR remains in "draft" state after failure (AC-06)

---

## TASK-005-04-07: Add loading indicator on Submit button during submission

| Field | Value |
|-------|-------|
| Target | `Views/Pages/PurchaseRequisitionPage.xaml` |
| Estimate | S |
| Depends On | TASK-005-04-01 |
| Blocks | None |

### What to do
- Add a small `ProgressRing` or spinner inside or next to the Submit button
- Bind its visibility to `{Binding IsSubmitting}`
- When `IsSubmitting == true`, show the spinner and optionally change button text to "Submitting..."
- When `IsSubmitting == false`, hide the spinner and restore button text to "Submit for Approval"
- The Submit button should also be disabled during submission (bound to `IsSubmitting` via CanExecute)

### How to verify
- [ ] Loading indicator visible during submission (AC-08)
- [ ] Button text or appearance changes during submission (AC-08)

---

## Dependency Graph

```
TASK-005-04-03 (SubmitPRAsync service)
       |
       v
TASK-005-04-02 (SubmitPRCommand) <-- TASK-005-04-01 (XAML button)
       |
       +--------> TASK-005-04-04 (confirmation dialog)
       +--------> TASK-005-04-05 (update state in UI)
       +--------> TASK-005-04-06 (error handling)

TASK-005-04-01 (XAML button) --> TASK-005-04-07 (loading indicator)
```
