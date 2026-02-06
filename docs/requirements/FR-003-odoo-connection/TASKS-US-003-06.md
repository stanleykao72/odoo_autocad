# TASKS: US-003-06 — Search Products

> **Parent US**: [US-003-06](US-003-06-search-products.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-05 (product sync and DataGrid) must be completed so products exist to search
- [ ] US-003-03 (connection status) must be completed for IsConnected gating

## Acceptance Criteria
- [ ] AC-01: A search box is visible above the product DataGrid
- [ ] AC-02: Entering text and triggering search filters products by name or code using `SearchProductsAsync(searchTerm)`
- [ ] AC-03: The search term must be at least 1 character before triggering a server-side search
- [ ] AC-04: Search results update the product DataGrid in real-time
- [ ] AC-05: The search box is disabled when `IsConnected` is false
- [ ] AC-06: An empty search result displays the message "No products found matching '{searchTerm}'."
- [ ] AC-07: The search operation runs asynchronously without blocking the UI thread
- [ ] AC-08: A "clear search" action restores the full product list from the local cache

---

## TASK-003-06-01: Add product search TextBox and button to XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-06-05 |

### What to do
- In the Product Catalog section toolbar (created in US-003-05), add or verify:
  - A `TextBox` with placeholder text "Search by name or code..." bound to `{Binding ProductSearchTerm, UpdateSourceTrigger=PropertyChanged}`
  - A "Search" `Button` bound to `{Binding SearchProductsCommand}`
  - A "Clear" `Button` (or X icon button) bound to `{Binding ClearProductSearchCommand}` to reset search
- Bind `TextBox.IsEnabled` and search `Button.IsEnabled` to `{Binding IsConnected}`
- Add a `TextBlock` for empty search results message bound to `{Binding ProductSearchEmptyMessage}` with `Visibility` bound to whether Products collection is empty after a search
- Optionally wire `TextBox.KeyDown` event to trigger search on Enter key press

### How to verify
- [ ] Search box is visible above the product DataGrid (AC-01)
- [ ] Clear button is available to reset the search (AC-08)

---

## TASK-003-06-02: Add ProductSearchTerm property to ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-06-03 |

### What to do
- Add observable property using `[ObservableProperty]`:
  - `private string _productSearchTerm = string.Empty;`
  - `private string _productSearchEmptyMessage = string.Empty;`
- Implement `partial void OnProductSearchTermChanged(string value)`:
  - Call `SearchProductsCommand.NotifyCanExecuteChanged()` to re-evaluate whether search is allowed
  - Optionally implement debounce: use a `DispatcherTimer` or `CancellationTokenSource` to delay the search by 300ms after the last keystroke, cancelling any pending search
  - If value is empty, call `ClearProductSearch()` to restore the full list
- Store a reference to the full cached product list so it can be restored on clear

### How to verify
- [ ] ProductSearchTerm property exists with change notification (AC-02)
- [ ] Debounce prevents excessive API calls during rapid typing (AC-07)

---

## TASK-003-06-03: Implement SearchProductsCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-06-02, TASK-003-06-04 |
| Blocks | None |

### What to do
- Add `SearchProductsCommand` as `IAsyncRelayCommand` using `[RelayCommand(CanExecute = nameof(CanSearchProducts))]` attribute on `async Task SearchProductsAsync()` method
- In `SearchProductsAsync()`:
  - Validate `ProductSearchTerm.Length >= 1` (VR-003-012)
  - Call `var results = await _odooService.SearchProductsAsync(ProductSearchTerm)`
  - Clear and repopulate `Products` collection with results
  - If `results.Count == 0`, set `ProductSearchEmptyMessage = $"No products found matching '{ProductSearchTerm}'."`
  - Else clear `ProductSearchEmptyMessage`
  - Update `TotalProductCount = results.Count`
- Add `bool CanSearchProducts() => IsConnected && !string.IsNullOrWhiteSpace(ProductSearchTerm)`
- Handle exceptions gracefully: display error in `ProductSearchEmptyMessage` and log
- Use `CancellationToken` to support cancellation if user changes search term while a search is in progress

### How to verify
- [ ] Search filters products by name or code using SearchProductsAsync (AC-02)
- [ ] Search results update the DataGrid (AC-04)
- [ ] Empty results display the correct message (AC-06)
- [ ] Search runs asynchronously (AC-07)

---

## TASK-003-06-04: Implement SearchProductsAsync in OdooService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-003-06-03 |

### What to do
- Verify the existing `SearchProductsAsync(string searchTerm)` implementation in `OdooService`:
  ```csharp
  var domain = new object[]
  {
      "|",
      new object[] { "name", "ilike", searchTerm },
      new object[] { "default_code", "ilike", searchTerm }
  };
  ```
- The existing implementation uses JSON-RPC `search_read` on `product.product` model with `ilike` operator for case-insensitive partial matching -- this is correct
- Add a `limit` parameter to prevent unbounded result sets: modify to use `SearchReadAsync` with `limit: 100` or make it configurable
- Add `CancellationToken` parameter support for cancellation during debounced searches
- Consider adding a `search_count` call first to report total matching records when results are truncated

### How to verify
- [ ] SearchProductsAsync queries by name or code with ilike operator (AC-02)
- [ ] Search term minimum 1 character is enforced before API call (AC-03)

---

## TASK-003-06-05: Add search input validation and IsEnabled binding

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-06-01 |
| Blocks | None |

### What to do
- Bind the search `TextBox.IsEnabled` to `{Binding IsConnected}` so it is disabled when not connected
- Bind the search `Button.IsEnabled` via the `SearchProductsCommand` binding (automatically disabled when `CanSearchProducts()` returns false)
- The `CanSearchProducts()` check already enforces `IsConnected && !string.IsNullOrWhiteSpace(ProductSearchTerm)`, so the minimum 1 character requirement (VR-003-012) is handled
- Add visual feedback: when `IsConnected` is false, show disabled styling on the search controls

### How to verify
- [ ] Search box is disabled when IsConnected is false (AC-05)
- [ ] Search requires minimum 1 character (AC-03)

---

## TASK-003-06-06: Implement clear search to restore cached product list

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add `ClearProductSearchCommand` as `IRelayCommand` using `[RelayCommand]` attribute on `void ClearProductSearch()` method
- In `ClearProductSearch()`:
  - Set `ProductSearchTerm = string.Empty`
  - Clear `ProductSearchEmptyMessage`
  - Restore `Products` collection from the locally cached full product list:
    - Maintain a private `_allProducts` field of type `List<OdooProduct>` that is populated during `SyncProductsAsync()`
    - Repopulate `Products.Clear()` followed by adding all items from `_allProducts`
  - Update `TotalProductCount` to `_allProducts.Count`
- If no cached products exist (never synced), leave `Products` empty
- Alternatively, if using `IMemoryCache`, retrieve the full list from cache key `"all_products"`

### How to verify
- [ ] Clear search restores the full product list from cache (AC-08)
- [ ] ProductSearchTerm is cleared (AC-08)

---

## Dependency Graph
```
TASK-003-06-04 (OdooService SearchProductsAsync)
    |
TASK-003-06-02 (ProductSearchTerm property)
    |
    +---> TASK-003-06-03 (SearchProductsCommand)

TASK-003-06-01 (XAML search box)
    |
    +---> TASK-003-06-05 (Validation and IsEnabled)

TASK-003-06-06 (Clear search / restore cache)
```
