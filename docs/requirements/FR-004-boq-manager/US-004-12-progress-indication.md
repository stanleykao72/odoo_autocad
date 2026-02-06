# US-004-12: Progress Indication

## User Story
**As a** CAD Engineer,
**I want to** see progress indication during extraction and push,
**So that** I know the system is working and not frozen.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: During extraction, a progress bar shows the percentage of layouts processed
- [ ] AC-02: During extraction, a text label shows the current layout being processed (e.g., "Layout 3/5 extracting...")
- [ ] AC-03: During push, a progress indicator shows records processed out of total (e.g., "Processing 15/42...")
- [ ] AC-04: The progress bar is visible only during active operations (`IsExtracting` or `IsPushing` is true)
- [ ] AC-05: The Extract and Push buttons show a "busy" state (disabled with spinner or changed text) during their respective operations
- [ ] AC-06: If an operation takes longer than expected, the progress text provides reassurance (e.g., "Still processing, please wait...")
- [ ] AC-07: A Cancel button is available to abort long-running operations

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-007 | Display progress indicator during extraction showing current layout being processed | Should |
| FR-004-021 | Display progress indicator during push showing records processed out of total | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-12-01 | Add ProgressBar XAML bound to ExtractionProgressPercent with BoolToVisibility on IsExtracting | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-12-02 | Add progress text label bound to ExtractionProgress (e.g., "Layout 3/5") | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-12-03 | Implement IProgress<T> reporting in ExtractCommand to update progress per layout | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-12-04 | Add push progress bar and text bound to IsPushing and push progress properties | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-12-05 | Implement IProgress<T> reporting in PushToOdooCommand to update progress per record | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-12-06 | Implement Cancel button with CancellationTokenSource for aborting long-running operations | `ViewModels/BOQManagerViewModel.cs` | M |
| TASK-004-12-07 | Add busy state styling for Extract and Push buttons during operations | `Views/Pages/BOQManagerPage.xaml` | S |

## Dependencies
- Depends on: US-004-01 (extraction has progress to report), US-004-04 (push has progress to report)
- Blocks: None

## Notes
- The extraction progress uses `ExtractionProgressPercent` (0.0 to 1.0) for the progress bar `Value` binding, and `ExtractionProgress` (string, e.g., "Layout 3/5") for the text label.
- Progress bar visibility is controlled by: `Visibility="{Binding IsExtracting, Converter={StaticResource BoolToVisibilityConverter}}"`.
- The `IProgress<T>` pattern is used to report progress from async commands back to the UI thread. The ViewModel creates a `Progress<(int current, int total, string message)>` and passes it to the extraction/push methods.
- The Cancel button uses a `CancellationTokenSource` that is passed to the async commands. When cancelled, the operation stops gracefully and reports partial results.
- During push, progress tracking is per-layout since the `import2boq_v2` API may process layouts in batch. The progress text should reflect the actual granularity of the API response.
- The `IsPushing` and `IsExtracting` boolean properties control not only progress bar visibility but also button enabled states (Extract and Push are disabled during their respective operations).
