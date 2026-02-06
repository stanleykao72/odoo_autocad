# US-003-11: Filter by Category

## User Story
**As a** CAD Engineer,
**I want to** filter products by category,
**So that** I only see products relevant to construction/engineering.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: A category filter dropdown (ComboBox) is visible alongside the product search box
- [ ] AC-02: The dropdown includes an "All" option to show all categories (null category filter)
- [ ] AC-03: Selecting a category filters products via `GetProductsByCategoryAsync(categoryId)`
- [ ] AC-04: The category filter works in combination with the text search
- [ ] AC-05: The category filter accepts null (all categories) or a valid positive integer category ID
- [ ] AC-06: The category filter dropdown is disabled when `IsConnected` is false
- [ ] AC-07: The filter operation runs asynchronously without blocking the UI thread
- [ ] AC-08: Category options are populated from Odoo's category tree (not hardcoded)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-014 | The page SHALL support filtering products by category via `GetProductsByCategoryAsync(categoryId)`, matching the Python filter of `categ_id child_of 27` for construction products | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-11-01 | Add category filter ComboBox alongside the product search TextBox in XAML | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-11-02 | Add SelectedCategoryId nullable int property to ViewModel | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-11-03 | Implement FilterByCategoryCommand as IAsyncRelayCommand calling IOdooService.GetProductsByCategoryAsync(categoryId) | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-11-04 | Implement GetProductsByCategoryAsync(categoryId) in OdooService using JSON-RPC with categ_id domain filter | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-11-05 | Fetch and populate category list from Odoo category tree for the ComboBox options | `OdooAutoCAD.Core/Odoo/OdooService.cs` | M |
| TASK-003-11-06 | Add category filter validation: null for all, positive integer for specific category (VR-003-014) | `ViewModels/OdooConnectionViewModel.cs` | S |

## Dependencies
- Depends on: US-003-05, US-003-03
- Blocks: None

## Notes
- The Python code hardcodes `categ_id child_of 27` for construction products. The C# version should make this configurable either via `appsettings.json` or via a UI dropdown populated from Odoo's category tree.
- Validation rule VR-003-014 applies: category filter accepts null (all categories) or a valid positive integer category ID.
- The wireframe shows "Category: [All v]" ComboBox next to the search field in the Product Catalog section.
- The `child_of` domain operator in Odoo returns all products in the specified category and its subcategories, which is useful for hierarchical category structures.
- Consider caching the category tree locally after the first fetch to avoid repeated API calls.
