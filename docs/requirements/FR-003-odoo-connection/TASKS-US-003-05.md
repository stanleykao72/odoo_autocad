# TASKS: US-003-05 — Sync Product Catalog

> **Parent US**: [US-003-05](US-003-05-sync-product-catalog.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 7 | **Effort**: 3S + 3M + 1L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form) must be completed
- [ ] US-003-03 (connection status and IsConnected property) must be completed

## Acceptance Criteria
- [ ] AC-01: A "Sync Products" button is visible on the Product Catalog section of the page
- [ ] AC-02: Clicking "Sync Products" calls `GetProductsAsync()` to fetch the full product catalog from Odoo
- [ ] AC-03: The button is disabled when `IsConnected` is false
- [ ] AC-04: The button is disabled while a sync operation is in progress (`IsSyncing` is true)
- [ ] AC-05: A progress indicator is shown during the sync operation without blocking the UI thread
- [ ] AC-06: On completion, a summary displays: total records fetched, timestamp of sync, and any errors encountered
- [ ] AC-07: Synced product data is cached locally (via SQLite or in-memory) to reduce repeated API calls within the same session
- [ ] AC-08: Products are displayed in a searchable, scrollable DataGrid with columns: ID, Name, Code, Description, Unit Price, UOM, Category
- [ ] AC-09: If sync partially fails, both the successful count and the list of errors are displayed
- [ ] AC-10: The DataGrid uses VirtualizingStackPanel for performance with large product catalogs

---

## TASK-003-05-01: Create Product Catalog section in XAML with DataGrid and controls

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | L |
| Depends On | None |
| Blocks | TASK-003-05-05 |

### What to do
- Add a `GroupBox` with `Header="Product Catalog"` in the lower portion of `OdooConnectionPage.xaml`
- Add a toolbar row with:
  - Search `TextBox` bound to `{Binding ProductSearchTerm}` with placeholder text "Search by name or code..."
  - Search `Button` bound to `{Binding SearchProductsCommand}`
  - Category filter `ComboBox` (placeholder for US-003-11)
  - "Sync Products" `Button` bound to `{Binding SyncProductsCommand}`
- Add a `DataGrid` with `AutoGenerateColumns="False"` and the following column definitions:
  - `DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="60"`
  - `DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"`
  - `DataGridTextColumn Header="Code" Binding="{Binding Code}" Width="100"`
  - `DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="200"`
  - `DataGridTextColumn Header="Unit Price" Binding="{Binding ListPrice, StringFormat=N2}" Width="100"`
  - `DataGridTextColumn Header="UOM" Binding="{Binding UnitOfMeasure}" Width="80"`
  - `DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="120"`
- Bind `DataGrid.ItemsSource` to `{Binding Products}`
- Add a `ProgressBar IsIndeterminate="True"` with `Visibility` bound to `{Binding IsSyncing, Converter={StaticResource BoolToVisibilityConverter}}`
- Add a summary `TextBlock` bound to `{Binding SyncSummaryText}` for showing record count and sync status
- Add `DataGrid.SelectedItem` bound to `{Binding SelectedProduct}`

### How to verify
- [ ] "Sync Products" button is visible (AC-01)
- [ ] DataGrid displays columns: ID, Name, Code, Description, Unit Price, UOM, Category (AC-08)
- [ ] Progress indicator is present and bound to IsSyncing (AC-05)

---

## TASK-003-05-02: Implement SyncProductsCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-003-05-06 |

### What to do
- Add `SyncProductsCommand` as `IAsyncRelayCommand` using `[RelayCommand(CanExecute = nameof(CanSyncProducts))]` attribute on `async Task SyncProductsAsync()` method
- In `SyncProductsAsync()`:
  - Set `IsSyncing = true`
  - Call `var products = await _odooService.GetProductsAsync()`
  - Populate `Products` `ObservableCollection<OdooProduct>` with the result
  - Set `TotalProductCount = products.Count`
  - Set `LastSyncTime = DateTime.Now`
  - Update `SyncSummaryText = $"Synced {products.Count} products at {DateTime.Now:HH:mm:ss}"`
  - Cache products locally (delegate to TASK-003-05-04)
  - Add entry to `SyncHistory` collection
  - Set `IsSyncing = false` in a `finally` block
- Add `bool CanSyncProducts() => IsConnected && !IsSyncing`
- Handle exceptions: on failure, display error in `SyncSummaryText` and log via `_logger`
- Call `SyncProductsCommand.NotifyCanExecuteChanged()` in `OnIsConnectedChanged` and `OnIsSyncingChanged`

### How to verify
- [ ] Clicking "Sync Products" calls GetProductsAsync (AC-02)
- [ ] Products populate the DataGrid (AC-08)
- [ ] Summary displays total records, timestamp, errors (AC-06)

---

## TASK-003-05-03: Add Products collection and sync status properties to ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-05-02 |

### What to do
- Add observable properties using `[ObservableProperty]`:
  - `private bool _isSyncing;`
  - `private int _totalProductCount;`
  - `private string _syncSummaryText = string.Empty;`
  - `private OdooProduct? _selectedProduct;`
- Add collection property (initialized in constructor):
  - `public ObservableCollection<OdooProduct> Products { get; } = new();`
- Implement `partial void OnIsSyncingChanged(bool value)` to:
  - Call `NotifyCanExecuteChanged()` on `SyncProductsCommand`, `SearchProductsCommand`, `FilterByCategoryCommand`, `SyncToOdooCommand`
- Validation rule VR-003-008: sync commands only available when `IsConnected`
- Validation rule VR-003-013: concurrent sync prevented by checking `IsSyncing`

### How to verify
- [ ] Button disabled when IsConnected is false (AC-03)
- [ ] Button disabled while sync is in progress (AC-04)
- [ ] UI thread is not blocked during sync (AC-05)

---

## TASK-003-05-04: Implement local product caching

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Data/Context/AppDbContext.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `ProductCache` entity to `OdooAutoCAD.Data.Entities`:
  ```csharp
  public class ProductCache
  {
      public int Id { get; set; }
      public int OdooProductId { get; set; }
      public string Name { get; set; } = string.Empty;
      public string? Code { get; set; }
      public string? Description { get; set; }
      public decimal? ListPrice { get; set; }
      public string? UnitOfMeasure { get; set; }
      public string? Category { get; set; }
      public DateTime CachedAt { get; set; } = DateTime.UtcNow;
  }
  ```
- Add `DbSet<ProductCache> ProductCaches { get; set; }` to `AppDbContext`
- Add `OnModelCreating` configuration for `ProductCache` with unique index on `OdooProductId`
- Create a `ProductCacheService` class (or add methods to an existing service) with:
  - `Task CacheProductsAsync(IEnumerable<OdooProduct> products)` -- clears old cache, inserts new
  - `Task<IReadOnlyList<OdooProduct>> GetCachedProductsAsync()` -- retrieves from SQLite
  - `Task<bool> HasCacheAsync()` -- returns true if cache exists and is recent
- Alternatively, use `IMemoryCache` from `Microsoft.Extensions.Caching.Memory` for in-session caching without SQLite persistence

### How to verify
- [ ] Synced products are cached locally (AC-07)
- [ ] Subsequent page loads can display cached products without re-fetching (AC-07)

---

## TASK-003-05-05: Configure DataGrid with VirtualizingStackPanel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | TASK-003-05-01 |
| Blocks | None |

### What to do
- Add virtualization settings to the product `DataGrid`:
  ```xml
  <DataGrid VirtualizingPanel.IsVirtualizing="True"
            VirtualizingPanel.VirtualizationMode="Recycling"
            VirtualizingPanel.ScrollUnit="Pixel"
            EnableRowVirtualization="True"
            EnableColumnVirtualization="True"
            MaxHeight="400"
            ScrollViewer.CanContentScroll="True">
  ```
- Set `DataGrid.RowDetailsVisibilityMode="Collapsed"` to avoid rendering row details for off-screen rows
- Consider adding `DataGrid.MaxHeight` or placing inside a `ScrollViewer` with fixed height to constrain the visible area
- Test with 1000+ rows to verify smooth scrolling performance

### How to verify
- [ ] DataGrid uses VirtualizingStackPanel for performance (AC-10)
- [ ] Large product catalogs scroll smoothly (AC-10)

---

## TASK-003-05-06: Implement sync result summary display

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-05-02 |
| Blocks | TASK-003-05-07 |

### What to do
- After successful sync in `SyncProductsAsync()`:
  - Format `SyncSummaryText = $"Synced {TotalProductCount} products successfully at {LastSyncTime:yyyy-MM-dd HH:mm}"`
- After partial failure:
  - Format `SyncSummaryText = $"Sync completed with errors: {successCount} succeeded, {failCount} failed out of {totalCount}"`
  - Store error details in a `SyncErrors` `ObservableCollection<string>` property for display in an expandable panel
- After complete failure:
  - Format `SyncSummaryText = $"Sync failed: {errorMessage}"`
- Add `LastSyncResult` property of type `SyncResult?` to store the full result for detailed inspection
- Display the summary in the XAML `TextBlock` with conditional foreground color (green=success, orange=partial, red=failure)

### How to verify
- [ ] Summary displays total records fetched, timestamp, errors (AC-06)
- [ ] Partial failures show both success and error counts (AC-09)

---

## TASK-003-05-07: Add SyncHistory entry after product sync

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-05-06 |
| Blocks | None |

### What to do
- Add `SyncHistory` property:
  - `public ObservableCollection<SyncHistoryEntry> SyncHistory { get; } = new();`
- After each sync operation (product sync, BOQ sync, parameter sync), insert a `SyncHistoryEntry`:
  ```csharp
  SyncHistory.Insert(0, new SyncHistoryEntry
  {
      SyncType = "Products",
      Timestamp = DateTime.Now,
      Success = result.Success,
      RecordsProcessed = result.RecordsProcessed,
      ErrorCount = result.RecordsFailed,
      Summary = SyncSummaryText
  });
  ```
- Insert at index 0 to show newest entries first (descending timestamp order)
- Optionally persist `SyncHistory` to `SyncLog` table in `AppDbContext` for cross-session history
- The sync history is displayed in the Sync Controls panel as shown in the wireframe

### How to verify
- [ ] Sync history entry is added after each sync operation (AC-06)
- [ ] History shows type, timestamp, status, and record count (AC-06)

---

## Dependency Graph
```
TASK-003-05-03 (ViewModel collections/properties)
    |
    +---> TASK-003-05-02 (SyncProductsCommand)
              |
              +---> TASK-003-05-06 (Sync result summary)
                        |
                        +---> TASK-003-05-07 (SyncHistory entry)

TASK-003-05-01 (XAML Product Catalog section)
    |
    +---> TASK-003-05-05 (VirtualizingStackPanel config)

TASK-003-05-04 (Local product caching)
```
