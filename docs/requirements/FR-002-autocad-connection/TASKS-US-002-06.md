# TASKS: US-002-06 — Clear Table IDs

> **Parent US**: [US-002-06](US-002-06-clear-table-ids.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 3S + 5M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-002-01 (AutoCAD connection must be established for write-back operations)

## Acceptance Criteria
- [ ] AC-01: The user can update block attributes in AutoCAD from the UI
- [ ] AC-02: A "Clear IDs (This Layout)" button clears the table Detail ID column for the currently selected layout
- [ ] AC-03: A "Clear All" button clears table Detail ID columns for ALL layouts in the drawing
- [ ] AC-04: Both clear operations display a confirmation dialog before proceeding (VR-002-006)
- [ ] AC-05: Upon completion, a success or failure message is displayed to the user
- [ ] AC-06: After clearing, the table data display is refreshed to reflect the cleared IDs
- [ ] AC-07: If an attribute write fails, an error message is displayed: "Failed to update attribute '{tag}' in AutoCAD."
- [ ] AC-08: Clear buttons are only enabled when AutoCAD is connected
- [ ] AC-09: All write-back COM operations execute on the GUI/STA thread via IGUIProxy

---

## TASK-002-06-01: Add "Clear IDs (This Layout)" and "Clear All" buttons to Actions Bar

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Actions Bar section at the bottom of `AutoCADPage.xaml` (alongside the "Extract Parameters" button), add:
  - A `Button` with Content "Clear IDs (This Layout)":
    - Bind `Command` to `ClearTableIdsCommand`
    - Bind `IsEnabled` via `CanExecute` to `IsConnected && SelectedLayout != null`
    - Apply a caution-styled appearance (e.g., orange border or icon)
  - A `Button` with Content "Clear All":
    - Bind `Command` to `ClearAllTableIdsCommand`
    - Bind `IsEnabled` via `CanExecute` to `IsConnected`
    - Apply a danger-styled appearance (e.g., red border or icon) to indicate broader impact
- Layout the three buttons in a horizontal `StackPanel` or `WrapPanel` with spacing:
  ```xml
  <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,8,0,0">
      <Button Content="Extract Parameters" Command="{Binding ExtractParametersCommand}" Margin="0,0,8,0"/>
      <Button Content="Clear IDs (This Layout)" Command="{Binding ClearTableIdsCommand}" Margin="0,0,8,0"/>
      <Button Content="Clear All" Command="{Binding ClearAllTableIdsCommand}"/>
  </StackPanel>
  ```

### How to verify
- [ ] Both clear buttons are visible in the Actions Bar (AC-02, AC-03)
- [ ] Buttons are disabled when AutoCAD is not connected (AC-08)

---

## TASK-002-06-02: Implement ClearTableIdsCommand (single layout) with confirmation dialog

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-002-06-07, TASK-002-06-08 |
| Blocks | None |

### What to do
- Add `[RelayCommand(CanExecute = nameof(CanClearTableIds))]` on:
  ```csharp
  private async Task ClearTableIdsAsync()
  ```
- Implement `bool CanClearTableIds()` returning `IsConnected && HasDocument && SelectedLayout != null`
- Implementation:
  - Show a confirmation dialog (via `IDialogService` or `MessageBox`):
    - Message: $"Clear all table IDs in layout '{SelectedLayout.Name}'? This operation cannot be undone."
    - Buttons: Yes/No
  - If user confirms:
    - Call `_guiProxy.ExecuteInGuiAsync("clear_table_id", new Dictionary<string, object?> { { "layoutName", SelectedLayout.Name } }, timeout: 30000)`
    - On success: set `StatusMessage = "Table IDs cleared successfully for layout '{name}'."`
    - On failure: set `ErrorMessage` with the error details
    - After completion: re-trigger `ExtractParametersAsync()` to refresh the table data display (AC-06)
  - If user cancels: do nothing
- VR-002-006: Confirmation dialog is mandatory before proceeding

### How to verify
- [ ] Confirmation dialog appears before clearing (AC-04)
- [ ] Success/failure message is displayed after operation (AC-05)
- [ ] Table data refreshes after clearing (AC-06)

---

## TASK-002-06-03: Implement ClearAllTableIdsCommand (all layouts) with confirmation dialog

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | M |
| Depends On | TASK-002-06-07, TASK-002-06-08 |
| Blocks | None |

### What to do
- Add `[RelayCommand(CanExecute = nameof(CanClearAllTableIds))]` on:
  ```csharp
  private async Task ClearAllTableIdsAsync()
  ```
- Implement `bool CanClearAllTableIds()` returning `IsConnected && HasDocument`
- Implementation:
  - Show a confirmation dialog with stronger warning:
    - Message: "Clear ALL table IDs in ALL layouts? This will affect every layout in the current drawing and cannot be undone."
    - Buttons: Yes/No
  - If user confirms:
    - Call `_guiProxy.ExecuteInGuiAsync("clear_all_tables_id", timeout: 60000)` (longer timeout for all layouts)
    - On success: set `StatusMessage = "Table IDs cleared for all layouts."`
    - On failure: set `ErrorMessage` with the error details
    - After completion: re-trigger `ExtractParametersAsync()` to refresh the table data display
  - If user cancels: do nothing
- VR-002-006: Confirmation dialog is mandatory before proceeding

### How to verify
- [ ] Confirmation dialog appears with strong warning before clearing all (AC-04)
- [ ] Success/failure message is displayed after operation (AC-05)
- [ ] Table data refreshes after clearing all (AC-06)

---

## TASK-002-06-04: Implement clear_table_id equivalent in AutoCADService for single layout

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | TASK-002-06-06 |
| Blocks | None |

### What to do
- Create a method in `AutoCADService`:
  ```csharp
  public bool ClearTableIds(string layoutName)
  ```
- Implementation:
  - Switch to the specified layout using `SwitchToLayout(layoutName)`
  - Get all tables in the layout (iterate PaperSpace entities, find `AcDbTable` entities)
  - For each valid table (9 columns, HEADER_ID in col 7):
    - Iterate data rows (row 2 to RowCount-1)
    - Set column 8 (Detail ID) to empty string using `SetTableValue(handle, row, 8, "")`
  - Return `true` if all operations succeed, `false` if any fail
- Register IGUIProxy handler:
  - `"clear_table_id"` action: takes `layoutName` parameter, calls `ClearTableIds(layoutName)`, returns bool
- Log each table processed and any errors encountered
- Handle COM exceptions with meaningful messages referencing the specific attribute that failed (AC-07)

### How to verify
- [ ] Detail ID column (index 8) is cleared for the specified layout (AC-02)
- [ ] COM operations go through IGUIProxy (AC-09)

---

## TASK-002-06-05: Implement clear_all_tables_id equivalent for all layouts

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | TASK-002-06-04 |
| Blocks | None |

### What to do
- Create a method in `AutoCADService`:
  ```csharp
  public (int succeeded, int failed) ClearAllTableIds()
  ```
- Implementation:
  - Get all layouts using `GetLayouts()`, filter out Model space
  - For each non-Model layout:
    - Call `ClearTableIds(layout.Name)`
    - Track success/failure count
  - Return tuple with count of succeeded and failed layouts
- Register IGUIProxy handler:
  - `"clear_all_tables_id"` action: calls `ClearAllTableIds()`, returns result object with counts
- Log progress: "Clearing table IDs: layout {n} of {total}..."
- If a single layout fails, continue with remaining layouts (do not abort entire operation)
- Return the original layout as active after processing (save and restore `GetCurrentLayoutName()`)

### How to verify
- [ ] Detail ID column is cleared for ALL layouts in the drawing (AC-03)
- [ ] Failures in individual layouts do not stop the entire operation

---

## TASK-002-06-06: Implement set_attribute_value equivalent for writing block attributes

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-06-04 |

### What to do
- Create a method in `AutoCADService` for writing attribute values back to AutoCAD:
  ```csharp
  public bool SetAttributeValue(string layoutName, string blockName, string tag, string value)
  ```
- Implementation:
  - Switch to the specified layout
  - Find the block reference by name in the layout's PaperSpace
  - Iterate `block.GetAttributes()` to find the attribute with matching `TagString`
  - Set `attr.TextString = value`
  - Return `true` on success
- Register IGUIProxy handler:
  - `"set_attribute_value"` action: takes `layoutName`, `blockName`, `tag`, `value` parameters
- Error handling:
  - If block not found, log error and return `false`
  - If attribute not found within block, log error and return `false`
  - Catch `COMException` and set error message: "Failed to update attribute '{tag}' in AutoCAD." (AC-07)
- This method is used by both clear operations and potential future attribute editing features (AC-01)

### How to verify
- [ ] Block attributes can be updated in AutoCAD from the application (AC-01)
- [ ] Error message displayed on write failure (AC-07)

---

## TASK-002-06-07: Define SetTableValue() and ClearSelection() in IAutoCADService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-06-02, TASK-002-06-03 |

### What to do
- Verify existing interface declarations:
  - `void SetTableValue(string tableHandle, int row, int column, string value)` -- already present (line 310)
  - `void ClearSelection()` -- already present (line 365)
- Add new method declarations if not already present:
  - `bool ClearTableIds(string layoutName)` -- clears Detail ID column for a single layout
  - `(int succeeded, int failed) ClearAllTableIds()` -- clears Detail IDs for all layouts
  - `bool SetAttributeValue(string layoutName, string blockName, string tag, string value)` -- writes an attribute value back to AutoCAD
- Add XML doc comments referencing FR-002-021 through FR-002-024
- Ensure method signatures match the implementations in TASK-002-06-04, 05, 06

### How to verify
- [ ] Interface declares ClearTableIds, ClearAllTableIds, SetAttributeValue methods
- [ ] Existing SetTableValue and ClearSelection are still present (AC-01)

---

## TASK-002-06-08: Add success/failure notification display after write-back operations

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-06-02, TASK-002-06-03 |

### What to do
- Implement a notification system in `AutoCADViewModel` for operation results:
  - Add `[ObservableProperty] private string _operationResultMessage = string.Empty;` -- displayed to user after operations
  - Add `[ObservableProperty] private bool _isOperationSuccess;` -- determines message styling (green for success, red for failure)
- Create a helper method:
  ```csharp
  private void ShowOperationResult(string message, bool isSuccess)
  {
      OperationResultMessage = message;
      IsOperationSuccess = isSuccess;
      // Optionally auto-clear after a delay
  }
  ```
- Use this for all write-back operations:
  - Clear single layout success: `"Table IDs cleared successfully for layout '{name}'."`
  - Clear all layouts success: `"Table IDs cleared for all layouts ({n} succeeded, {m} failed)."`
  - Failure: `"Failed to update attribute '{tag}' in AutoCAD."` or generic error message
- In XAML: bind a notification bar/banner to `OperationResultMessage`:
  - Use `Visibility` converter on `OperationResultMessage` being non-empty
  - Use `IsOperationSuccess` to switch between success (green) and error (red) styling

### How to verify
- [ ] Success message displayed after successful clear operation (AC-05)
- [ ] Failure message displayed after failed write-back (AC-05, AC-07)

---

## Dependency Graph
```
TASK-002-06-07 (IAutoCADService interface additions)
    |
    +---> TASK-002-06-02 (ClearTableIdsCommand)
    |
    +---> TASK-002-06-03 (ClearAllTableIdsCommand)

TASK-002-06-06 (SetAttributeValue service)
    |
    +---> TASK-002-06-04 (ClearTableIds single layout)
              |
              +---> TASK-002-06-05 (ClearAllTableIds all layouts)

TASK-002-06-08 (Notification display)
    |
    +---> TASK-002-06-02 (ClearTableIdsCommand)
    |
    +---> TASK-002-06-03 (ClearAllTableIdsCommand)

TASK-002-06-01 (XAML buttons)
```
