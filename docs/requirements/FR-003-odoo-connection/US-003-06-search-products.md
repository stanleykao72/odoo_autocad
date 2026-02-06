# US-003-06: Search Products

## User Story
**As a** CAD Engineer,
**I want to** search and filter products by name or code,
**So that** I can quickly find the product I need.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A search box is visible above the product DataGrid
- [ ] AC-02: Entering text and triggering search filters products by name or code using `SearchProductsAsync(searchTerm)`
- [ ] AC-03: The search term must be at least 1 character before triggering a server-side search
- [ ] AC-04: Search results update the product DataGrid in real-time
- [ ] AC-05: The search box is disabled when `IsConnected` is false
- [ ] AC-06: An empty search result displays the message "No products found matching '{searchTerm}'."
- [ ] AC-07: The search operation runs asynchronously without blocking the UI thread
- [ ] AC-08: A "clear search" action restores the full product list from the local cache

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-013 | The page SHALL provide a search box that filters products by name or code in real-time using `SearchProductsAsync(searchTerm)` | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-06-01 | Add product search TextBox with search button above the product DataGrid in XAML | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-06-02 | Add ProductSearchTerm property to ViewModel with change notification | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-06-03 | Implement SearchProductsCommand as IAsyncRelayCommand calling IOdooService.SearchProductsAsync(searchTerm) | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-06-04 | Implement SearchProductsAsync(searchTerm) in OdooService using JSON-RPC search_read with name/code domain | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-06-05 | Add input validation: minimum 1 character, bind search IsEnabled to IsConnected and validation state | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-06-06 | Implement clear search functionality to restore full cached product list | `ViewModels/OdooConnectionViewModel.cs` | S |

## Dependencies
- Depends on: US-003-05, US-003-03
- Blocks: None

## Notes
- Validation rule VR-003-012 requires the search term to be at least 1 character before triggering a server-side search to prevent unbounded queries.
- The Python equivalent is `search_products(query, limit=50)` which searches products with a text query and result limit.
- Consider implementing debounce logic on the search input to avoid excessive API calls during rapid typing.
- If products are cached locally, the search can first filter the local cache and fall back to server-side search for terms not found locally.
