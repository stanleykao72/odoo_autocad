# TASKS: US-002-02 — View Layouts

> **Parent US**: [US-002-02](US-002-02-view-layouts.md)
> **Parent FR**: [FR-002](FR-002-autocad-connection.md)
> **Priority**: P1
> **Tasks**: 7 | **Effort**: 3S + 3M + 0L
> **Status**: Done

## Prerequisites
- [x] US-002-01 (AutoCAD connection must be established first)

## Acceptance Criteria
- [x] AC-01: When connected to AutoCAD, the page displays a list of all layouts in the current drawing, excluding "Model"
- [x] AC-02: Each layout entry shows the layout name and tab order
- [x] AC-03: The currently active layout is visually highlighted in the list
- [x] AC-04: The user can select a layout from the list to view its details in the right panel
- [x] AC-05: The user can switch the active layout in AutoCAD by selecting a different layout
- [x] AC-06: The layout list is refreshed when the connection is first established and can be manually refreshed
- [x] AC-07: All layout-related COM operations execute on the GUI/STA thread via IGUIProxy

---

## TASK-002-02-01: Add Layouts ListView panel to the AutoCAD page XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-02-06, TASK-002-02-07 |

### What to do
- Add a left panel section below the Connection Status panel in `AutoCADPage.xaml`
- Create a `ListView` or `ListBox` bound to `Layouts` (`ObservableCollection<LayoutInfo>`)
- Set `SelectedItem` binding to `SelectedLayout` with `Mode=TwoWay`
- Define an `ItemTemplate` (`DataTemplate`) showing:
  - `TextBlock` bound to `Name` for layout name
  - `TextBlock` bound to `TabOrder` for tab order number
  - Layout: use a horizontal `StackPanel` or `Grid` with two columns
- Use a `Grid` with two columns for the layout area: left (Layouts list, ~30% width) and right (Layout Details panel, ~70% width)
- Bind `Visibility` or `IsEnabled` to `IsConnected` so the list is disabled when disconnected (VR-002-002)
- Add a "Refresh Layouts" button below the list bound to a refresh command

### How to verify
- [x] Layouts ListView is visible and shows layout name + tab order per entry (AC-01, AC-02)
- [x] ListView is disabled when AutoCAD is not connected (VR-002-002)

---

## TASK-002-02-02: Implement Layouts ObservableCollection and SelectedLayout property in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-002-02-06 |

### What to do
- Add to `AutoCADViewModel`:
  - `public ObservableCollection<LayoutInfo> Layouts { get; } = new();` -- collection of layouts
  - `[ObservableProperty] private LayoutInfo? _selectedLayout;` -- currently selected layout
  - `[ObservableProperty] private string? _activeLayoutName;` -- name of the active layout in AutoCAD
- Implement `partial void OnSelectedLayoutChanged(LayoutInfo? value)` to trigger `SelectLayoutCommand` execution when a layout is selected
- Add a method `LoadLayoutsAsync()` that:
  - Calls `_guiProxy.ExecuteInGuiAsync("get_layouts")` to retrieve layouts
  - Clears `Layouts` collection and repopulates from result
  - Filters out entries where `IsModelSpace == true`
  - Sets `ActiveLayoutName` from `_guiProxy.ExecuteInGuiAsync("get_active_layout")`

### How to verify
- [x] Layouts collection is populated and excludes "Model" (AC-01)
- [x] SelectedLayout property triggers detail loading on change (AC-04)

---

## TASK-002-02-03: Implement GetLayouts() in AutoCADService excluding "Model"

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-002-02-04 |

### What to do
- The existing `GetLayouts()` method already iterates `_acadDoc.Layouts` and returns `IReadOnlyList<LayoutInfo>`
- Verify it correctly:
  - Iterates all layouts via `foreach (dynamic layout in _acadDoc.Layouts)`
  - Creates `LayoutInfo` records with `Name`, `TabOrder`, `IsModelSpace`, `PlotConfigurationName`
  - Orders by `TabOrder`
- The "Model" filtering is done at the ViewModel level (LayoutInfo has `IsModelSpace` flag), but confirm `IsModelSpace` is set correctly: `layout.Name == "Model"`
- Ensure this method handles `_acadDoc == null` gracefully by returning empty list
- Register handler with IGUIProxy: `"get_layouts"` action that calls `GetLayouts()` and returns result

### How to verify
- [x] GetLayouts returns all layouts including Model flag, ordered by TabOrder (AC-01, AC-02)
- [x] Returns empty list when no document is open

---

## TASK-002-02-04: Implement GetActiveLayout() and SetActiveLayout(name) in AutoCADService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` |
| Estimate | M |
| Depends On | TASK-002-02-03 |
| Blocks | None |

### What to do
- The existing `GetCurrentLayoutName()` returns `_acadDoc?.ActiveLayout?.Name` -- verify this works correctly
- The existing `SwitchToLayout(string layoutName)` sets `_acadDoc.ActiveLayout = _acadDoc.Layouts.Item(layoutName)` -- verify
- Register IGUIProxy handlers:
  - `"get_active_layout"` action: calls `GetCurrentLayoutName()`, returns layout name string
  - `"set_active_layout"` action: takes `layoutName` from parameters dict, calls `SwitchToLayout(layoutName)`, returns bool success
- Ensure both methods log operations and handle exceptions with user-friendly messages
- Add validation: if `layoutName` is null/empty or doesn't exist in the drawing, return false with appropriate log message

### How to verify
- [x] GetCurrentLayoutName returns the active layout name (AC-03)
- [x] SwitchToLayout changes the active layout in AutoCAD (AC-05)

---

## TASK-002-02-05: Define GetLayouts(), GetActiveLayout() in IAutoCADService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify existing interface declarations:
  - `IReadOnlyList<LayoutInfo> GetLayouts()` -- already present
  - `bool SwitchToLayout(string layoutName)` -- already present
  - `string? GetCurrentLayoutName()` -- already present
- These are all confirmed in the existing interface skeleton (lines 152-165)
- No new methods needed; verify XML doc comments reference FR-002-008 through FR-002-012
- Ensure `LayoutInfo` record includes all necessary fields: `Name`, `TabOrder`, `IsModelSpace`, `PlotConfigurationName` (already defined)

### How to verify
- [x] Interface declares GetLayouts, SwitchToLayout, GetCurrentLayoutName (AC-01, AC-05)

---

## TASK-002-02-06: Implement SelectLayoutCommand to load layout details on selection change

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` |
| Estimate | S |
| Depends On | TASK-002-02-01, TASK-002-02-02 |
| Blocks | None |

### What to do
- Implement `SelectLayoutCommand` as `IRelayCommand<LayoutInfo>` using `[RelayCommand]` on a method:
  ```csharp
  private async Task SelectLayout(LayoutInfo? layout)
  ```
- When a layout is selected:
  - Call `_guiProxy.ExecuteInGuiAsync("set_active_layout", new Dictionary<string, object?> { { "layoutName", layout.Name } })`
  - Update `ActiveLayoutName` to the selected layout name
  - Trigger layout details loading (this feeds into US-002-03 parameter extraction)
  - Log the selection change
- Set `CanExecute` to require `IsConnected == true` and `layout != null`
- Wire `OnSelectedLayoutChanged` partial method to invoke this command

### How to verify
- [x] Selecting a layout switches the active layout in AutoCAD (AC-04, AC-05)
- [x] Command is disabled when not connected

---

## TASK-002-02-07: Add visual highlighting for the active layout in the ListView

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` |
| Estimate | S |
| Depends On | TASK-002-02-01 |
| Blocks | None |

### What to do
- Add a `DataTrigger` or `Style` to the `ListView.ItemContainerStyle` that highlights the active layout:
  - Bind a `MultiBinding` or use a converter comparing `LayoutInfo.Name` with `DataContext.ActiveLayoutName`
  - When matched, set `Background` to a highlight color (e.g., `#E3F2FD` light blue) and `FontWeight` to `Bold`
- Alternatively, add an `IsActive` computed property in a wrapper ViewModel class or use an `IValueConverter`:
  - `ActiveLayoutConverter` implementing `IMultiValueConverter` that compares layout name with active layout name
- Ensure the highlight updates when `ActiveLayoutName` changes (PropertyChanged notification)
- The active layout should be visually distinct from the selected (clicked) item

### How to verify
- [x] The currently active layout is visually highlighted in the list (AC-03)
- [x] Highlight updates when the active layout changes (AC-06)

---

## Dependency Graph
```
TASK-002-02-05 (IAutoCADService interface verification)

TASK-002-02-03 (GetLayouts service impl)
    |
    +---> TASK-002-02-04 (GetActive/SetActive layout)

TASK-002-02-01 (XAML Layouts ListView)
    |
    +---> TASK-002-02-07 (Active layout highlighting)
    |
    +---> TASK-002-02-06 (SelectLayoutCommand)
              ^
              |
TASK-002-02-02 (Layouts collection & SelectedLayout)
```
