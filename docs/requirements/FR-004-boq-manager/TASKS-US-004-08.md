# TASKS: US-004-08 — Map Products

> **Parent US**: [US-004-08](US-004-08-map-products.md)
> **Parent FR**: [FR-004](FR-004-boq-manager.md)
> **Priority**: P1
> **Tasks**: 8 | **Effort**: 4S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-004-03 (Validation identifies unrecognized products — errors with field "ProductId" must exist)
- [ ] US-003-01 (Odoo connected — `IOdooService.SearchProductsAsync()` requires active connection)

## Acceptance Criteria
- [ ] AC-01: A Product Mapping dialog is accessible from the validation results panel when an unrecognized product is flagged
- [ ] AC-02: The dialog shows the unrecognized AutoCAD `product_no` value and provides a search field to query Odoo products
- [ ] AC-03: Product mapping checks the local mapping cache (`IBOQProcessor.GetProductMappings()`) first before querying Odoo via `IOdooService.SearchProductsAsync()`
- [ ] AC-04: Product matching is case-insensitive as implemented by `StringComparer.OrdinalIgnoreCase` dictionary
- [ ] AC-05: The search field queries Odoo products and displays results for user selection
- [ ] AC-06: User-selected mappings are persisted via `IBOQProcessor.SetProductMapping()` and available within the same application lifecycle
- [ ] AC-07: After a mapping is applied, the affected BOQEntryRow is updated with `MappedOdooProductId` and `MappedOdooProductName`, and `IsProductMapped` is set to true
- [ ] AC-08: After mapping, re-validation of the affected entry clears the product-related validation error

---

## TASK-004-08-01: Create ProductMappingDialog XAML with AutoCAD name display, search field, results list, and Apply/Cancel buttons

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Dialogs/ProductMappingDialog.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-08-06 |

### What to do
- Create a new WPF `Window` in namespace `OdooAutoCAD.App.Views.Dialogs` with `x:Class="OdooAutoCAD.App.Views.Dialogs.ProductMappingDialog"`
- Set `WindowStartupLocation="CenterOwner"`, `SizeToContent="WidthAndHeight"`, `MinWidth="450"`, `MinHeight="350"`, `Title="Product Mapping"`
- Layout sections:
  1. **AutoCAD Name** (read-only): `TextBlock` with bold text bound to `{Binding AutoCADProductName}`
  2. **Search**: `TextBox` bound to `{Binding SearchTerm}` with a "Search" `Button` bound to `{Binding SearchCommand}`
  3. **Results**: `ListView` with `RadioButton` items bound to `{Binding SearchResults}` (`ObservableCollection<OdooProduct>`), showing `Name`, `Code`, `Category`
  4. **Selected**: display selected product info bound to `{Binding SelectedProduct}`
  5. **Buttons**: "Apply Mapping" bound to `{Binding ApplyCommand}` (sets `DialogResult=true`) and "Cancel" bound to close
- Add `ProductMappingDialog.xaml.cs` code-behind that sets `DataContext` to `ProductMappingDialogViewModel`
- Set `Owner` to the main window when opening the dialog

### How to verify
- [ ] Dialog shows AutoCAD product name (read-only), search field, results list, Apply/Cancel buttons (AC-01, AC-02)
- [ ] Results list shows Odoo product details with radio selection (AC-05)

---

## TASK-004-08-02: Implement ProductMappingDialogViewModel with search command and Odoo product query

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/Dialogs/ProductMappingDialogViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-08-06 |

### What to do
- Create `ProductMappingDialogViewModel` in namespace `OdooAutoCAD.App.ViewModels.Dialogs`, extending `ObservableObject`
- Inject `IOdooService` and `IBOQProcessor` via constructor
- Properties: `string AutoCADProductName` (set from caller), `string SearchTerm`, `ObservableCollection<OdooProduct> SearchResults`, `OdooProduct? SelectedProduct`
- `SearchCommand` (`IAsyncRelayCommand`): calls `_odooService.SearchProductsAsync(SearchTerm)` and populates `SearchResults`
- `ApplyCommand` (`IRelayCommand`): validates `SelectedProduct != null`, calls `_boqProcessor.SetProductMapping(AutoCADProductName, SelectedProduct.Id)`, returns result
- On initialization, check if a mapping already exists via `_boqProcessor.GetProductMappings()` and pre-populate if found
- Handle empty search results with a "No products found" message

### How to verify
- [ ] Search queries Odoo and populates results list (AC-05)
- [ ] Apply calls SetProductMapping with the selected product (AC-06)
- [ ] Local cache is checked before Odoo query (AC-03)

---

## TASK-004-08-03: Implement GetProductMappings() in IBOQProcessor returning case-insensitive dictionary

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify the existing `GetProductMappings()` implementation in `BOQProcessor` returns the `_productMappings` dictionary
- Confirm the dictionary is initialized with `StringComparer.OrdinalIgnoreCase` (already present in the skeleton: `new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)`)
- Ensure the return type `IReadOnlyDictionary<string, int>` is correct for read-only access from consumers
- Add a `ContainsMapping(string autocadName)` convenience method that checks `_productMappings.ContainsKey(autocadName)` (case-insensitive)

### How to verify
- [ ] GetProductMappings returns case-insensitive dictionary (AC-04)
- [ ] Existing mappings are accessible from consumers (AC-03)

---

## TASK-004-08-04: Implement SetProductMapping() in IBOQProcessor for persisting user-defined mappings

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/BOQ/BOQProcessor.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Verify the existing `SetProductMapping(string autocadName, int odooProductId)` implementation in `BOQProcessor`
- Ensure it adds/updates the entry in `_productMappings[autocadName] = odooProductId`
- The case-insensitive comparer handles key normalization automatically
- Add logging: `_logger.LogInformation("Product mapping added: '{AutoCADName}' -> Odoo ID {ProductId}", autocadName, odooProductId)`
- Mappings persist in-memory for the application lifecycle (no database persistence in this iteration)

### How to verify
- [ ] SetProductMapping persists the mapping in the in-memory dictionary (AC-06)
- [ ] Mapping is available for subsequent validation and push operations (AC-06)

---

## TASK-004-08-05: Implement SearchProductsAsync() in IOdooService for querying Odoo product catalog

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-004-08-02 |

### What to do
- Implement `SearchProductsAsync(string searchTerm)` in `OdooService`
- Build an Odoo RPC search query filtering products by name/code containing the search term (case-insensitive)
- Filter by `category_id = 27` and `active = true` (matching the Python reference `get_product()` behavior)
- Map Odoo response records to `OdooProduct` records with `Id`, `Name`, `Code`, `Description`, `ListPrice`, `UnitOfMeasure`, `Category`
- Return `IReadOnlyList<OdooProduct>` — empty list if no results found
- Implement pagination if needed (Odoo default limit is typically 80 records)
- Add timeout handling and logging

### How to verify
- [ ] Search queries Odoo for products matching the search term (AC-05)
- [ ] Results include Id, Name, Code, Category for display in the dialog (AC-05)

---

## TASK-004-08-06: Wire "Map Product" action button in validation results panel to open ProductMappingDialog

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-08-01, TASK-004-08-02 |
| Blocks | TASK-004-08-07 |

### What to do
- Define `OpenProductMappingCommand` as `IRelayCommand<BOQEntryRow>` in `BOQManagerViewModel`
- The command creates a `ProductMappingDialogViewModel` with `AutoCADProductName = entry.ProductNo`
- Instantiates `ProductMappingDialog`, sets its `DataContext` and `Owner`, calls `ShowDialog()`
- If `DialogResult == true`, retrieves the selected `OdooProduct` from the dialog ViewModel
- This command is invoked from the "Map Product" button in the validation results panel (bound via `CommandParameter="{Binding}"` on the error item's data context)
- Also wire the `OpenProductMappingCommand` from the DataGrid context menu for direct row-level mapping

### How to verify
- [ ] "Map Product" action opens the ProductMappingDialog with the correct AutoCAD product name (AC-01)
- [ ] Dialog result is captured after user selects and applies mapping (AC-01)

---

## TASK-004-08-07: Update BOQEntryRow mapping state after successful product mapping

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-08-06 |
| Blocks | TASK-004-08-08 |

### What to do
- After `ProductMappingDialog` returns with a successful mapping:
  - Set `entry.MappedOdooProductId = selectedProduct.Id`
  - Set `entry.MappedOdooProductName = selectedProduct.Name`
  - Set `entry.IsProductMapped = true`
- Also find and update all other `BOQEntryRow` items in `AllEntries` that have the same `ProductNo` (since the mapping applies globally)
- Notify property changes on each updated entry to refresh the DataGrid display
- Update the `ProductMappings` `ObservableCollection` to include the new mapping for the mapping management view

### How to verify
- [ ] Affected BOQEntryRow is updated with MappedOdooProductId, MappedOdooProductName, IsProductMapped (AC-07)
- [ ] All entries with the same ProductNo are updated (AC-07)

---

## TASK-004-08-08: Trigger re-validation of affected entries after mapping is applied

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` |
| Estimate | S |
| Depends On | TASK-004-08-07 |
| Blocks | None |

### What to do
- After mapping state is updated on the affected entries, trigger re-validation:
  - Option A: Call `ExecuteValidateAsync()` to re-validate all entries
  - Option B: Perform targeted re-validation only on affected entries (more efficient) by calling `_boqProcessor.ValidateBOQAsync()` with just the affected subset
- After re-validation, the product-related `BOQValidationError` for the mapped product should be cleared (since it now resolves to a valid Odoo product)
- Remove the corresponding entry from `ValidationErrors` collection
- Update `HasValidationErrors` and summary counts via `UpdateSummaryCounts()`
- If all errors are resolved, `CanPush` becomes true

### How to verify
- [ ] After mapping, re-validation clears the product-related error (AC-08)
- [ ] Summary counts and HasValidationErrors are updated (AC-08)

---

## Dependency Graph
```
TASK-004-08-05 (SearchProductsAsync Service)
       │
       └──▶ TASK-004-08-02 (Dialog ViewModel)

TASK-004-08-01 (Dialog XAML) ──────────┐
TASK-004-08-02 (Dialog ViewModel) ─────┤
                                       ▼
                            TASK-004-08-06 (Wire Map Product Action)
                                       │
                                       └──▶ TASK-004-08-07 (Update Entry State)
                                                  │
                                                  └──▶ TASK-004-08-08 (Re-validation)

TASK-004-08-03 (GetProductMappings) ─── independent
TASK-004-08-04 (SetProductMapping) ─── independent

Tasks 01, 03, 04, 05 can begin in parallel.
Task 02 depends on 05. Task 06 depends on 01 and 02.
Tasks 07 and 08 are sequential after 06.
```
