# TASKS: US-004-12 — Progress Indication

> **Parent US**: [US-004-12](US-004-12-progress-indication.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P2
> **Tasks**: 7 | **Effort**: 4S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-01 (Extraction has progress to report — extraction pipeline must support `IProgress<T>`)
- [ ] US-004-04 (Push has progress to report — push pipeline must support `IProgress<T>`)

## Acceptance Criteria
- [ ] AC-01: During extraction, a progress bar shows the percentage of layouts processed
- [ ] AC-02: During extraction, a text label shows the current layout being processed (e.g., "Layout 3/5 extracting...")
- [ ] AC-03: During push, a progress indicator shows records processed out of total (e.g., "Processing 15/42...")
- [ ] AC-04: The progress bar is visible only during active operations (`IsExtracting` or `IsPushing` is true)
- [ ] AC-05: The Extract and Push buttons show a "busy" state (disabled with spinner or changed text) during their respective operations
- [ ] AC-06: If an operation takes longer than expected, the progress text provides reassurance (e.g., "Still processing, please wait...")
- [ ] AC-07: A Cancel button is available to abort long-running operations

---

## TASK-004-12-01: Add ProgressBar XAML bound to ExtractionProgressPercent with BoolToVisibility on IsExtracting

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a WPF `ProgressBar` in the extraction panel section of `BOQManagerPage.xaml`
- Bind `Value="{Binding ExtractionProgressPercent}"` with `Maximum="1.0"` and `Minimum="0.0"`
- Bind `Visibility="{Binding IsExtracting, Converter={StaticResource BoolToVisibilityConverter}}"` so the bar is only visible during extraction
- Set `Height="6"` for a slim progress bar, with a blue or accent color `Foreground`
- Position below the extraction buttons and above the summary bar
- When `IsExtracting` is false, the ProgressBar collapses and takes no space

### How to verify
- [ ] Progress bar shows percentage during extraction (AC-01)
- [ ] Progress bar is hidden when not extracting (AC-04)

---

## TASK-004-12-02: Add progress text label bound to ExtractionProgress

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `TextBlock` next to or below the extraction `ProgressBar`
- Bind `Text="{Binding ExtractionProgress}"` — the ViewModel sets this to strings like "Layout 3/5 extracting..."
- Bind `Visibility="{Binding IsExtracting, Converter={StaticResource BoolToVisibilityConverter}}"` to match the ProgressBar visibility
- Style with a subtle gray color and smaller font to differentiate from primary content
- The text updates in real-time as each layout is processed

### How to verify
- [ ] Text label shows "Layout X/Y extracting..." during extraction (AC-02)
- [ ] Text is hidden when not extracting (AC-04)

---

## TASK-004-12-03: Implement IProgress<T> reporting in ExtractCommand to update progress per layout

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- In `ExecuteExtractAsync()`, create a `Progress<(int current, int total, string layoutName)>` instance
- Register a callback that updates ViewModel properties on each report:
  - `ExtractionProgressPercent = (double)current / total`
  - `ExtractionProgress = $"Layout {current}/{total} extracting {layoutName}..."`
- Pass the progress reporter to the extraction pipeline (via `IGUIProxy.ExecuteInGuiAsync()` parameters or as part of the extraction handler)
- If extraction takes longer than 10 seconds on a single layout, update the text to "Still processing {layoutName}, please wait..."
- Use a `Stopwatch` to track per-layout duration and trigger the reassurance message
- Initialize progress at 0% before extraction starts and set to 100% when complete
- Clear `ExtractionProgress` text after extraction completes

### How to verify
- [ ] Progress updates per layout with correct current/total values (AC-01, AC-02)
- [ ] Long-running layouts show reassurance message (AC-06)

---

## TASK-004-12-04: Add push progress bar and text bound to IsPushing and push progress properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `ProgressBar` in the push controls section bound to a `PushProgressPercent` property
- Bind `Visibility="{Binding IsPushing, Converter={StaticResource BoolToVisibilityConverter}}"`
- Add a `TextBlock` bound to `PushProgressText` (e.g., "Processing 15/42...")
- Position below the Push button and above the push result summary
- Add `[ObservableProperty] private double _pushProgressPercent;` and `[ObservableProperty] private string _pushProgressText = string.Empty;` to the ViewModel
- Same visibility pattern as extraction progress: hidden when not pushing

### How to verify
- [ ] Push progress bar shows during push operation (AC-03)
- [ ] Push progress text shows records processed out of total (AC-03)
- [ ] Progress is hidden when not pushing (AC-04)

---

## TASK-004-12-05: Implement IProgress<T> reporting in PushToOdooCommand to update progress per record

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- In `ExecutePushAsync()`, create a `Progress<(int current, int total, string message)>` instance
- Register a callback that updates:
  - `PushProgressPercent = (double)current / total`
  - `PushProgressText = $"Processing {current}/{total}..."`
- Since the `import2boq_v2` API may process layouts in batch, track progress at the layout level:
  - Report after each layout's data is submitted
  - Total = number of layouts; current = layouts processed so far
- If a single layout push takes longer than 15 seconds, update text to "Still processing, please wait..."
- Initialize at 0% before push starts, set to 100% when complete
- Clear progress text after push completes (replaced by `LastPushResult`)

### How to verify
- [ ] Push progress updates per layout batch (AC-03)
- [ ] Long operations show reassurance message (AC-06)

---

## TASK-004-12-06: Implement Cancel button with CancellationTokenSource for aborting long-running operations

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `CancelCommand` (`IRelayCommand`) to `BOQManagerViewModel`
- Maintain a `CancellationTokenSource? _currentCts` field
- In `ExecuteExtractAsync()` and `ExecutePushAsync()`:
  - Create a new `_currentCts = new CancellationTokenSource()`
  - Pass `_currentCts.Token` to the async operations
  - Check `token.IsCancellationRequested` between layouts/records and throw `OperationCanceledException` if cancelled
- `CancelCommand.Execute()`: calls `_currentCts?.Cancel()`
- On cancellation: set progress text to "Operation cancelled. Partial results may be available."
  - For extraction: show partial extracted data
  - For push: show partial push results (items already pushed to Odoo are preserved)
- Add a "Cancel" `Button` in `BOQManagerPage.xaml` visible during active operations:
  - `Visibility="{Binding IsAnyOperationInProgress, Converter={StaticResource BoolToVisibilityConverter}}"`
  - `IsAnyOperationInProgress` is `IsExtracting || IsPushing`
- Dispose `_currentCts` in the `finally` block of each operation

### How to verify
- [ ] Cancel button is visible during extraction and push operations (AC-07)
- [ ] Clicking Cancel aborts the operation gracefully (AC-07)
- [ ] Partial results are preserved after cancellation (AC-07)

---

## TASK-004-12-07: Add busy state styling for Extract and Push buttons during operations

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Apply visual busy state to the "Extract from AutoCAD" button when `IsExtracting` is true:
  - Change `Content` to "Extracting..." using a `DataTrigger` on `IsExtracting`
  - Disable the button (already handled by `CanExecute`)
  - Optionally add a small spinner animation using a `ProgressBar IsIndeterminate="True"` with `Width="16" Height="16"` inside the button content
- Apply the same pattern to "Push to Odoo" button when `IsPushing` is true:
  - Change `Content` to "Pushing..."
  - Disable the button
- Use `Style.Triggers` with `DataTrigger` bound to the boolean properties
- Also apply busy state to "Validate All" when `IsValidating` is true: content changes to "Validating..."

### How to verify
- [ ] Extract button shows "Extracting..." and disabled state during extraction (AC-05)
- [ ] Push button shows "Pushing..." and disabled state during push (AC-05)
- [ ] Buttons return to normal state after operations complete (AC-05)

---

## Dependency Graph
```
TASK-004-12-01 (Extraction ProgressBar XAML) ─── independent
TASK-004-12-02 (Extraction Progress Text XAML) ─── independent
TASK-004-12-03 (Extraction IProgress<T>) ─── independent
TASK-004-12-04 (Push Progress XAML) ─── independent
TASK-004-12-05 (Push IProgress<T>) ─── independent
TASK-004-12-06 (Cancel Button + CTS) ─── independent
TASK-004-12-07 (Busy State Styling) ─── independent

All 7 tasks are independent and can be developed in parallel.
XAML tasks (01, 02, 04, 07) are front-end only.
ViewModel tasks (03, 05, 06) are back-end logic.
No within-US dependencies exist; all tasks contribute independently to the progress indication feature.
```
