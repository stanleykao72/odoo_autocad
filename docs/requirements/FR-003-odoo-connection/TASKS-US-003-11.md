# TASKS: US-003-11 — Filter by Category

> **Parent US**: [US-003-11](US-003-11-filter-by-category.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P2
> **Tasks**: 6 | **Effort**: 2S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-05 (product sync and DataGrid) must be completed so products exist to filter
- [ ] US-003-03 (connection status) must be completed for IsConnected gating

## Acceptance Criteria
- [ ] AC-01: A category filter dropdown (ComboBox) is visible alongside the product search box
- [ ] AC-02: The dropdown includes an "All" option to show all categories (null category filter)
- [ ] AC-03: Selecting a category filters products via `GetProductsByCategoryAsync(categoryId)`
- [ ] AC-04: The category filter works in combination with the text search
- [ ] AC-05: The category filter accepts null (all categories) or a valid positive integer category ID
- [ ] AC-06: The category filter dropdown is disabled when `IsConnected` is false
- [ ] AC-07: The filter operation runs asynchronously without blocking the UI thread
- [ ] AC-08: Category options are populated from Odoo's category tree (not hardcoded)

---

## TASK-003-11-01: Add category filter ComboBox to XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Product Catalog toolbar (alongside the search TextBox), add a `ComboBox`:
  ```xml
  <ComboBox ItemsSource="{Binding Categories}"
            SelectedItem="{Binding SelectedCategory}"
            DisplayMemberPath="Name"
            IsEnabled="{Binding IsConnected}"
            Width="150"
            Margin="5,0">
  </ComboBox>
  ```
- The ComboBox should be positioned after the search TextBox and before the "Sync Products" button, matching the wireframe layout: "Category: [All v]"
- Add a label "Category:" before the ComboBox
- The "All" option should be the default selection, represented by a category entry with `Id = null` or a sentinel value

### How to verify
- [ ] Category filter ComboBox is visible alongside the product search box (AC-01)
- [ ] ComboBox is disabled when IsConnected is false (AC-06)

---

## TASK-003-11-02: Add SelectedCategoryId and Categories properties

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-11-03 |

### What to do
- Define a `CategoryItem` class for ComboBox display:
  ```csharp
  public class CategoryItem
  {
      public int? Id { get; set; }
      public string Name { get; set; } = string.Empty;
  }
  ```
- Add observable properties using `[ObservableProperty]`:
  - `private int? _selectedCategoryId;`
  - `private CategoryItem? _selectedCategory;`
- Add collection property:
  - `public ObservableCollection<CategoryItem> Categories { get; } = new();`
- Implement `partial void OnSelectedCategoryChanged(CategoryItem? value)`:
  - Set `SelectedCategoryId = value?.Id`
  - Call `FilterByCategoryCommand.NotifyCanExecuteChanged()`
  - Optionally auto-trigger the filter on selection change
- Initialize `Categories` with a default "All" entry:
  ```csharp
  Categories.Add(new CategoryItem { Id = null, Name = "All" });
  SelectedCategory = Categories[0];
  ```

### How to verify
- [ ] SelectedCategoryId property accepts null for all categories (AC-02, AC-05)
- [ ] Categories collection is populated for ComboBox binding (AC-01)

---

## TASK-003-11-03: Implement FilterByCategoryCommand

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-11-02 |
| Blocks | None |

### What to do
- Add `FilterByCategoryCommand` as `IAsyncRelayCommand` using `[RelayCommand(CanExecute = nameof(CanFilterByCategory))]` attribute on `async Task FilterByCategoryAsync()` method
- In `FilterByCategoryAsync()`:
  - If `SelectedCategoryId` is `null`, fetch all products: `await _odooService.GetProductsAsync()`
  - If `SelectedCategoryId` has a value, fetch filtered: `await _odooService.GetProductsByCategoryAsync(SelectedCategoryId.Value)`
  - If `ProductSearchTerm` is not empty, apply text search on top of category filter (combine with local filtering or make a combined API call)
  - Clear and repopulate `Products` collection with results
  - Update `TotalProductCount`
- Add `bool CanFilterByCategory() => IsConnected && !IsSyncing`
- Handle the combination of text search + category filter:
  - Option A: Fetch by category from server, then filter locally by search term
  - Option B: Build a combined domain with both `categ_id` and `name`/`code` ilike conditions
- Validation rule VR-003-014: category filter accepts null or valid positive integer

### How to verify
- [ ] Selecting a category filters products via GetProductsByCategoryAsync (AC-03)
- [ ] Category filter works in combination with text search (AC-04)
- [ ] Filter runs asynchronously (AC-07)

---

## TASK-003-11-04: Implement GetProductsByCategoryAsync with child_of operator

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-003-11-03 |

### What to do
- Modify the existing `GetProductsByCategoryAsync(int categoryId)` to use the `child_of` domain operator instead of `=`:
  ```csharp
  public async Task<IReadOnlyList<OdooProduct>> GetProductsByCategoryAsync(int categoryId)
  {
      var domain = new object[] { new object[] { "categ_id", "child_of", categoryId } };
      var result = await SearchReadAsync("product.product", domain,
          new[] { "id", "name", "default_code", "description", "list_price", "uom_id", "categ_id" });
      return result.Select(r => MapToOdooProduct(r)).ToList();
  }
  ```
- The `child_of` operator returns all products in the specified category AND its subcategories, matching the Python behavior of `categ_id child_of 27`
- This is more useful for hierarchical category trees than the current `=` operator
- Add `CancellationToken` support for the async operation
- Consider also adding an `ActiveOnly` filter: `new object[] { "active", "=", true }` to match the Python filter

### How to verify
- [ ] GetProductsByCategoryAsync uses child_of operator for hierarchical categories (AC-03)
- [ ] Returns products in the category and its subcategories (AC-03)

---

## TASK-003-11-05: Fetch and populate category list from Odoo

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | L |
| Depends On | None |
| Blocks | None |

### What to do
- Add a new method to `IOdooService`:
  ```csharp
  Task<IReadOnlyList<OdooCategory>> GetCategoriesAsync();
  ```
- Define the `OdooCategory` record:
  ```csharp
  public record OdooCategory(int Id, string Name, int? ParentId, string CompleteName);
  ```
- Implement in `OdooService`:
  ```csharp
  public async Task<IReadOnlyList<OdooCategory>> GetCategoriesAsync()
  {
      var result = await SearchReadAsync("product.category", Array.Empty<object>(),
          new[] { "id", "name", "parent_id", "complete_name" });
      return result.Select(r => new OdooCategory(
          Id: r.GetProperty("id").GetInt32(),
          Name: r.GetProperty("name").GetString() ?? "",
          ParentId: r.TryGetProperty("parent_id", out var pid) && pid.ValueKind == JsonValueKind.Array
              ? pid[0].GetInt32() : null,
          CompleteName: r.TryGetProperty("complete_name", out var cn) && cn.ValueKind == JsonValueKind.String
              ? cn.GetString() ?? "" : ""
      )).ToList();
  }
  ```
- In the ViewModel, populate `Categories` after successful connection:
  ```csharp
  var categories = await _odooService.GetCategoriesAsync();
  Categories.Clear();
  Categories.Add(new CategoryItem { Id = null, Name = "All" });
  foreach (var cat in categories)
      Categories.Add(new CategoryItem { Id = cat.Id, Name = cat.CompleteName });
  SelectedCategory = Categories[0];
  ```
- Cache the category list locally to avoid repeated API calls
- Consider showing only top-level categories or using a hierarchical tree display

### How to verify
- [ ] Category options are populated from Odoo's category tree, not hardcoded (AC-08)
- [ ] "All" option is available as default (AC-02)

---

## TASK-003-11-06: Add category filter validation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Validate `SelectedCategoryId` per VR-003-014:
  - Accept `null` (all categories -- no filter applied)
  - Accept positive integers (valid Odoo category IDs)
  - Reject negative numbers or zero
- In `CanFilterByCategory()`:
  ```csharp
  bool CanFilterByCategory() =>
      IsConnected &&
      !IsSyncing &&
      (SelectedCategoryId == null || SelectedCategoryId > 0);
  ```
- If a manually entered category ID is invalid, display a validation error
- Since the ComboBox is populated from the server, invalid IDs should not normally be possible -- but validate defensively
- Handle the edge case where a cached category ID no longer exists on the server (category was deleted): display a "Category not found" message and fallback to "All"

### How to verify
- [ ] Category filter accepts null for all categories (AC-05)
- [ ] Category filter accepts valid positive integer category IDs (AC-05)
- [ ] Invalid category values are rejected (AC-05)

---

## Dependency Graph
```
TASK-003-11-02 (ViewModel properties)
    |
    +---> TASK-003-11-03 (FilterByCategoryCommand)

TASK-003-11-04 (OdooService child_of operator)
    |
    +---> TASK-003-11-03

TASK-003-11-05 (Fetch category list from Odoo)

TASK-003-11-01 (XAML ComboBox)

TASK-003-11-06 (Category validation)
```
