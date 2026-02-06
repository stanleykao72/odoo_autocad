# TASKS: US-002-04 — Drawing Info

> **Parent US**: [US-002-04](US-002-04-drawing-info.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P1
> **Tasks**: 5 | **Effort**: 5S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-002-01 (AutoCAD connection must be established to read document info)

## Acceptance Criteria
- [ ] AC-01: Upon successful connection to AutoCAD, the page displays the current document filename (e.g., "Frame-001.dwg")
- [ ] AC-02: The full file path of the current document is displayed (e.g., "C:\Projects\Steel\Frame-001.dwg")
- [ ] AC-03: Document info is displayed in the Connection Status panel at the top of the page
- [ ] AC-04: Document info updates when the active document changes (on manual refresh)
- [ ] AC-05: When no document is open, a message is displayed: "No drawing is open in AutoCAD. Please open a drawing."
- [ ] AC-06: When no document is open, layout and table controls are disabled

---

## TASK-002-04-01: Add DocumentName and DocumentPath display fields to Connection Status panel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Connection Status panel of `AutoCADPage.xaml` (created in TASK-002-01-01), add:
  - A `TextBlock` with label "Document:" followed by a `TextBlock` bound to `DocumentName`
  - A `TextBlock` with label "Path:" followed by a `TextBlock` bound to `DocumentPath` (use `TextTrimming="CharacterEllipsis"` for long paths)
  - Use a `StackPanel` or `Grid` row to position these below the connection status indicator
- Add a conditional display for the "no document" message:
  - A `TextBlock` with text "No drawing is open in AutoCAD. Please open a drawing." bound via `Visibility` converter to `HasDocument` (visible when `HasDocument == false` and `IsConnected == true`)
- Style the document info text with a monospace or semi-bold font to distinguish file names

### How to verify
- [ ] Document filename is displayed in the Connection Status panel (AC-01, AC-03)
- [ ] Full file path is displayed (AC-02, AC-03)
- [ ] "No drawing" message appears when no document is open (AC-05)

---

## TASK-002-04-02: Implement DocumentName and DocumentPath properties with change notification

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-04-04, TASK-002-04-05 |

### What to do
- Verify the following `[ObservableProperty]` fields exist in `AutoCADViewModel` (added in TASK-002-01-06):
  - `private string _documentName = string.Empty;`
  - `private string _documentPath = string.Empty;`
- Add a computed property:
  ```csharp
  public bool HasDocument => !string.IsNullOrEmpty(DocumentName);
  ```
- Implement `partial void OnDocumentNameChanged(string value)` to:
  - Call `OnPropertyChanged(nameof(HasDocument))` to update the computed property
- These properties should be populated during the connection flow in `ConnectAsync()`:
  - After successful connection, extract from `AutoCADStatus.CurrentDocument` for the filename
  - Extract full path from `GetCurrentDocumentPath()` via IGUIProxy call

### How to verify
- [ ] DocumentName and DocumentPath are observable and update the UI on change (AC-01, AC-02)
- [ ] HasDocument returns false when no document is open

---

## TASK-002-04-03: Populate document info from ActiveDocument during connection

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify the existing `AutoCADService` connection flow populates document info:
  - After obtaining `_acadDoc = _acadApp.ActiveDocument`, extract:
    - `_acadDoc.Name` -- document filename (e.g., "Frame-001.dwg")
    - `_acadDoc.FullName` -- full file path (e.g., "C:\Projects\Steel\Frame-001.dwg")
  - These values are returned as part of `AutoCADStatus` via `GetStatusAsync()`
- Verify `GetCurrentDocumentPath()` returns `_acadDoc?.FullName` (already implemented)
- Ensure `GetStatusAsync()` populates `CurrentDocument` field with the document name
- Register IGUIProxy handler `"get_document_info"` that returns a dictionary with `name` and `path` keys:
  ```csharp
  { "name", _acadDoc?.Name }, { "path", _acadDoc?.FullName }
  ```
- Handle null `_acadDoc` gracefully -- return null values, do not throw

### How to verify
- [ ] Document name and path are available after connection (AC-01, AC-02)
- [ ] Null document is handled gracefully without exceptions (AC-05)

---

## TASK-002-04-04: Implement RefreshStatusCommand to update document info on demand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | TASK-002-04-02 |
| Blocks | None |

### What to do
- Implement `RefreshStatusCommand` using `[RelayCommand]` attribute:
  ```csharp
  [RelayCommand]
  private async Task RefreshStatusAsync()
  ```
- In `RefreshStatusAsync()`:
  - Call `_guiProxy.ExecuteInGuiAsync("get_autocad_status")` to get current `AutoCADStatus`
  - Update `Status` property with the returned `AutoCADStatus`
  - Call `_guiProxy.ExecuteInGuiAsync("get_document_info")` to refresh document info
  - Update `DocumentName` and `DocumentPath` from the response
  - If `IsConnected` but document info returns null, set the "no document" state
  - Update `IsConnected` based on the status response
- This command is bound to the "Refresh Status" button in the Connection Status panel
- Command should be enabled at all times (useful even when disconnected to re-check)

### How to verify
- [ ] Refresh button updates document info when active document changes (AC-04)
- [ ] Refresh correctly detects when no document is open (AC-05)

---

## TASK-002-04-05: Add "no document open" state handling with control disabling

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | TASK-002-04-02 |
| Blocks | None |

### What to do
- Implement state management for the "no document open" scenario:
  - When `IsConnected == true` but `HasDocument == false`:
    - Set `ErrorMessage = "No drawing is open in AutoCAD. Please open a drawing."`
    - Disable layout-related commands: `SelectLayoutCommand`, `ExtractParametersCommand`
    - Disable table-related commands: `ClearTableIdsCommand`, `ClearAllTableIdsCommand`
    - Clear `Layouts` collection, `TableRows` collection, `LayoutAttributes` dictionary
  - When `HasDocument` changes back to `true`:
    - Clear `ErrorMessage`
    - Re-enable layout and table commands
    - Trigger layout loading via `LoadLayoutsAsync()`
- Update `CanExecute` methods for dependent commands to include `HasDocument` check:
  ```csharp
  private bool CanExtractParameters() => IsConnected && HasDocument && SelectedLayout != null;
  ```
- Implement `partial void OnDocumentNameChanged(string value)` to trigger `CanExecute` re-evaluation on all dependent commands

### How to verify
- [ ] "No drawing" message is displayed when connected but no document open (AC-05)
- [ ] Layout and table controls are disabled when no document is open (AC-06)

---

## Dependency Graph
```
TASK-002-04-01 (XAML document display fields)

TASK-002-04-03 (Service document info population)

TASK-002-04-02 (ViewModel DocumentName/Path properties)
    |
    +---> TASK-002-04-04 (RefreshStatusCommand)
    |
    +---> TASK-002-04-05 (No document state handling)
```
