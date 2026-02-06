# TASKS: US-004-11 — Manage Mappings

> **Parent US**: [US-004-11](US-004-11-manage-mappings.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P3
> **Tasks**: 7 | **Effort**: 4S + 2M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-08 (Product mapping infrastructure must exist — `IBOQProcessor.GetProductMappings()` and `SetProductMapping()` implemented)

## Acceptance Criteria
- [ ] AC-01: A product mappings management view is accessible from the BOQ Manager page (e.g., via a "Manage Mappings" button or menu)
- [ ] AC-02: The management view displays all current product mappings as a list showing: AutoCAD Name, Odoo Product ID, and Odoo Product Name
- [ ] AC-03: The admin can add new mappings by specifying an AutoCAD product name and selecting an Odoo product
- [ ] AC-04: The admin can delete existing mappings
- [ ] AC-05: The admin can modify existing mappings by re-selecting the Odoo product for an AutoCAD name
- [ ] AC-06: Changes to mappings take effect immediately for subsequent validation and push operations
- [ ] AC-07: Mappings are persisted via `IBOQProcessor.SetProductMapping()` and available across extraction sessions within the same application lifecycle

---

## TASK-004-11-01: Add "Manage Mappings" button to BOQ Manager page that opens mapping management view

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Manage Mappings" `Button` in the BOQ Manager page — either in the extraction panel area or as a toolbar button
- Bind `Command="{Binding ManageMappingsCommand}"`
- Style with a settings/gear icon or neutral appearance (non-destructive action)
- Position near the validation/push controls area since mappings relate to product resolution
- The button should always be enabled (mappings can be managed even without extracted data)

### How to verify
- [ ] "Manage Mappings" button renders and is accessible from the BOQ Manager page (AC-01)

---

## TASK-004-11-02: Create mapping management section or dialog with DataGrid showing all current mappings

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Dialogs/ProductMappingDialog.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-11-04 |

### What to do
- Extend the existing `ProductMappingDialog.xaml` with a "Management" tab or create a separate management section
- Add a `DataGrid` showing all current mappings with columns:
  - "AutoCAD Name" bound to `{Binding AutoCADName}` (read-only)
  - "Odoo Product ID" bound to `{Binding OdooProductId}` (read-only)
  - "Odoo Product Name" bound to `{Binding OdooProductName}` (read-only)
- Add a toolbar above the DataGrid with buttons: "Add New", "Edit Selected", "Delete Selected"
- The DataGrid should be read-only; editing is done via the Add/Edit dialog flow
- Set `SelectionMode="Single"` for row selection
- Show "No mappings defined" placeholder when the list is empty

### How to verify
- [ ] Management view displays all current mappings with AutoCAD Name, Odoo Product ID, Odoo Product Name (AC-02)
- [ ] Management view includes Add, Edit, Delete action buttons (AC-03, AC-04, AC-05)

---

## TASK-004-11-03: Implement mapping management ViewModel with Add, Delete, and Modify commands

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/Dialogs/ProductMappingDialogViewModel.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-004-11-04 |

### What to do
- Add management commands to `ProductMappingDialogViewModel`:
  - `AddMappingCommand` (`IAsyncRelayCommand`): opens a sub-dialog or inline form to enter AutoCAD name and search/select Odoo product, then calls `_boqProcessor.SetProductMapping(name, productId)`
  - `DeleteMappingCommand` (`IRelayCommand<ProductMappingEntry>`): removes the selected mapping — since `IBOQProcessor` only has `SetProductMapping`, implement a `RemoveProductMapping(string autocadName)` method or directly modify the dictionary
  - `EditMappingCommand` (`IRelayCommand<ProductMappingEntry>`): opens the search flow pre-populated with the current mapping's AutoCAD name, allowing the user to re-select the Odoo product
- Properties: `ObservableCollection<ProductMappingEntry> AllMappings`, `ProductMappingEntry? SelectedMapping`
- `RefreshMappings()` method that reads from `_boqProcessor.GetProductMappings()` and populates `AllMappings` by resolving Odoo product names via `_odooService.GetProductAsync(productId)`
- Handle the case where an Odoo product is no longer available (show "Unknown Product" with the ID)

### How to verify
- [ ] Add command creates a new mapping via IBOQProcessor.SetProductMapping() (AC-03)
- [ ] Delete command removes the selected mapping (AC-04)
- [ ] Edit command allows re-selecting the Odoo product for an existing mapping (AC-05)

---

## TASK-004-11-04: Bind ProductMappings ObservableCollection to the management view DataGrid

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-11-02, TASK-004-11-03 |
| Blocks | None |

### What to do
- Define `ManageMappingsCommand` as `IRelayCommand` in `BOQManagerViewModel`
- When executed, create `ProductMappingDialogViewModel` in management mode (not single-product mapping mode)
- Call `RefreshMappings()` to load current mappings before showing the dialog
- Open the `ProductMappingDialog` with `Owner` set to the main window
- After the dialog closes, refresh `ProductMappings` `ObservableCollection` in the ViewModel to reflect any changes
- Trigger re-validation if mappings were modified (since mapping changes may resolve or introduce validation errors)

### How to verify
- [ ] Opening the management view shows current mappings loaded from IBOQProcessor (AC-02)
- [ ] Changes are reflected in the ViewModel after dialog closes (AC-06)

---

## TASK-004-11-05: Implement GetProductMappings() to return all current mappings for display

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify the existing `GetProductMappings()` returns `IReadOnlyDictionary<string, int>` which provides all current AutoCAD-to-Odoo product ID mappings
- The consumer (ViewModel) will iterate this dictionary and resolve each `int` product ID to an `OdooProduct` name via `IOdooService.GetProductAsync(id)` for display
- Ensure thread-safety: if mappings can be modified concurrently, consider returning a snapshot copy
- The case-insensitive `StringComparer.OrdinalIgnoreCase` ensures consistent key behavior

### How to verify
- [ ] GetProductMappings returns all current mappings as a read-only dictionary (AC-02)
- [ ] Thread-safety is handled for concurrent access scenarios

---

## TASK-004-11-06: Implement Add/Remove mapping operations through IBOQProcessor interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-004-11-03 |

### What to do
- Add a `RemoveProductMapping(string autocadName)` method to `IBOQProcessor` interface and implement in `BOQProcessor`:
  - `_productMappings.Remove(autocadName)` — returns `true` if removed, `false` if not found
  - Log: `_logger.LogInformation("Product mapping removed: '{AutoCADName}'", autocadName)`
- Verify `SetProductMapping()` handles both add and update scenarios (already implemented: `_productMappings[autocadName] = odooProductId`)
- Consider adding `ClearAllMappings()` for bulk reset if needed
- Update `IBOQProcessor` interface to include `bool RemoveProductMapping(string autocadName)`

### How to verify
- [ ] RemoveProductMapping successfully removes a mapping from the dictionary (AC-04)
- [ ] SetProductMapping handles both add and update (AC-03, AC-05)

---

## TASK-004-11-07: Wire Odoo product search for adding/modifying mappings

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/Dialogs/ProductMappingDialogViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Ensure the Add and Edit flows in the management ViewModel reuse the same Odoo product search functionality from the single-mapping dialog:
  - `SearchCommand` queries `_odooService.SearchProductsAsync(searchTerm)` and populates `SearchResults`
  - User selects a product from the results
  - Apply button saves the mapping
- For the Add flow: provide a `TextBox` for the AutoCAD name input (editable) plus the search/select flow
- For the Edit flow: the AutoCAD name is pre-populated (read-only) and only the Odoo product selection changes
- Changes take effect immediately since `SetProductMapping()` modifies the in-memory dictionary

### How to verify
- [ ] Odoo product search works for both add and edit flows (AC-03, AC-05)
- [ ] Changes take effect immediately for subsequent operations (AC-06)
- [ ] Mappings persist within the application lifecycle (AC-07)

---

## Dependency Graph
```
TASK-004-11-06 (Add/Remove Interface) ─────┐
                                            │
                                            └──▶ TASK-004-11-03 (Management ViewModel)

TASK-004-11-02 (Management XAML) ──────────┐
TASK-004-11-03 (Management ViewModel) ─────┤
                                            ▼
                                TASK-004-11-04 (Wire to BOQManagerViewModel)

TASK-004-11-01 (Manage Mappings Button) ─── independent
TASK-004-11-05 (GetProductMappings Verify) ─── independent
TASK-004-11-07 (Odoo Search Reuse) ─── independent

Tasks 01, 02, 05, 06, 07 can begin in parallel.
Task 03 depends on 06 (interface must support remove).
Task 04 depends on 02 and 03.
```
