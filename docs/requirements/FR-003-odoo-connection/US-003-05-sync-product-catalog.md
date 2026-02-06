# US-003-05: Sync Product Catalog

## User Story
**As a** CAD Engineer,
**I want to** synchronize the product catalog from Odoo,
**So that** I have an up-to-date list of products for BOQ entries.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

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

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-011 | The page SHALL provide a "Sync Products" button that fetches the full product catalog from Odoo via `GetProductsAsync()` | Must |
| FR-003-015 | Product sync results SHALL display a summary: total records fetched, timestamp of sync, and any errors encountered | Should |
| FR-003-016 | Synced product data SHALL be cached locally (via SQLite or in-memory) to reduce repeated API calls within the same session | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-05-01 | Create Product Catalog section in XAML with "Sync Products" button, DataGrid, and summary panel | `Views/Pages/OdooConnectionPage.xaml` | L |
| TASK-003-05-02 | Implement SyncProductsCommand as IAsyncRelayCommand calling IOdooService.GetProductsAsync() | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-05-03 | Add Products ObservableCollection, TotalProductCount, and IsSyncing properties to ViewModel | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-05-04 | Implement local product caching using AppDbContext or IMemoryCache to store fetched products | `OdooAutoCAD.Data/Context/AppDbContext.cs` | M |
| TASK-003-05-05 | Configure DataGrid with VirtualizingStackPanel for performance with large datasets | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-05-06 | Implement sync result summary display showing records fetched, timestamp, and error details | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-05-07 | Add SyncHistory entry after each product sync operation | `ViewModels/OdooConnectionViewModel.cs` | S |

## Dependencies
- Depends on: US-003-01, US-003-03
- Blocks: US-003-06, US-003-11

## Notes
- The Python version caches products to SQLite via SQLAlchemy. The C# version should cache to the local SQLite database (`Database:ConnectionString` in config) or use `IMemoryCache`.
- The Python `get_product()` uses a hardcoded domain filter `[('categ_id', 'child_of', 27), ('active', '=', True)]` for construction products. The full sync here fetches all products; category filtering is handled by US-003-11.
- Validation rule VR-003-008 ensures product sync is only available when `IsConnected` is true.
- Validation rule VR-003-013 prevents concurrent sync operations by disabling buttons while `IsSyncing` is true.
- The product DataGrid should show columns matching the `OdooProduct` record: Id, Name, Code, Description, ListPrice, UnitOfMeasure, Category.
