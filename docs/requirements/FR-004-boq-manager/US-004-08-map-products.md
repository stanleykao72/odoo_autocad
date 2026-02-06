# US-004-08: Map Products

## User Story
**As a** CAD Engineer,
**I want to** map unrecognized AutoCAD product names to Odoo products,
**So that** the push succeeds even when naming conventions differ.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: A Product Mapping dialog is accessible from the validation results panel when an unrecognized product is flagged
- [ ] AC-02: The dialog shows the unrecognized AutoCAD `product_no` value and provides a search field to query Odoo products
- [ ] AC-03: Product mapping checks the local mapping cache (`IBOQProcessor.GetProductMappings()`) first before querying Odoo via `IOdooService.SearchProductsAsync()`
- [ ] AC-04: Product matching is case-insensitive as implemented by `StringComparer.OrdinalIgnoreCase` dictionary
- [ ] AC-05: The search field queries Odoo products and displays results for user selection
- [ ] AC-06: User-selected mappings are persisted via `IBOQProcessor.SetProductMapping()` and available within the same application lifecycle
- [ ] AC-07: After a mapping is applied, the affected BOQEntryRow is updated with `MappedOdooProductId` and `MappedOdooProductName`, and `IsProductMapped` is set to true
- [ ] AC-08: After mapping, re-validation of the affected entry clears the product-related validation error

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-023 | Provide Product Mapping panel/dialog for mapping unrecognized product_no values to Odoo product IDs | Must |
| FR-004-024 | Check local mapping cache first before querying Odoo | Must |
| FR-004-025 | Support case-insensitive matching via StringComparer.OrdinalIgnoreCase | Must |
| FR-004-026 | Provide search field in mapping dialog that queries Odoo products for selection | Should |
| FR-004-027 | Persist user-defined mappings via IBOQProcessor.SetProductMapping() across extraction sessions | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-08-01 | Create ProductMappingDialog XAML with AutoCAD name display, search field, results list, and Apply/Cancel buttons | `Views/Dialogs/ProductMappingDialog.xaml` | M |
| TASK-004-08-02 | Implement ProductMappingDialogViewModel with search command and Odoo product query | `ViewModels/Dialogs/ProductMappingDialogViewModel.cs` | M |
| TASK-004-08-03 | Implement GetProductMappings() in IBOQProcessor returning case-insensitive dictionary | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-08-04 | Implement SetProductMapping() in IBOQProcessor for persisting user-defined mappings | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-08-05 | Implement SearchProductsAsync() in IOdooService for querying Odoo product catalog | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | M |
| TASK-004-08-06 | Wire "Map Product" action button in validation results panel to open ProductMappingDialog | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-08-07 | Update BOQEntryRow mapping state after successful product mapping | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-08-08 | Trigger re-validation of affected entries after mapping is applied | `ViewModels/BOQManagerViewModel.cs` | S |

## Dependencies
- Depends on: US-004-03 (validation identifies unrecognized products)
- Blocks: US-004-04 (push requires all products to be resolved)

## Notes
- The Product Mapping dialog is a modal `Window` opened from the validation results panel. It contains:
  - The unrecognized AutoCAD `product_no` value (read-only display)
  - A search text field that queries Odoo products
  - A list of search results with radio button selection
  - Apply and Cancel buttons
- Mapping resolution follows a two-tier approach: `IBOQProcessor.GetProductMappings()` (in-memory `Dictionary<string, int>` with `OrdinalIgnoreCase`) is checked first. Only if no match is found does the system query `IOdooService.SearchProductsAsync()`.
- The C# implementation adds this local mapping cache that does not exist in the Python reference, where product resolution is handled entirely by Odoo.
- Mappings persist for the application lifecycle (in-memory). Persistent storage across sessions could be a future enhancement.
