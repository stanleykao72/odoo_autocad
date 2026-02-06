# US-004-11: Manage Mappings

## User Story
**As a** System Admin,
**I want to** view product mapping rules and modify them,
**So that** consistent mapping is maintained across sessions.

## Parent Feature
- **FR**: [FR-004-boq-manager](../FR-004-boq-manager/FR-004-boq-manager.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: A product mappings management view is accessible from the BOQ Manager page (e.g., via a "Manage Mappings" button or menu)
- [ ] AC-02: The management view displays all current product mappings as a list showing: AutoCAD Name, Odoo Product ID, and Odoo Product Name
- [ ] AC-03: The admin can add new mappings by specifying an AutoCAD product name and selecting an Odoo product
- [ ] AC-04: The admin can delete existing mappings
- [ ] AC-05: The admin can modify existing mappings by re-selecting the Odoo product for an AutoCAD name
- [ ] AC-06: Changes to mappings take effect immediately for subsequent validation and push operations
- [ ] AC-07: Mappings are persisted via `IBOQProcessor.SetProductMapping()` and available across extraction sessions within the same application lifecycle

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-004-027 | User-defined product mappings persist via IBOQProcessor.SetProductMapping() and are available across extraction sessions | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-004-11-01 | Add "Manage Mappings" button to BOQ Manager page that opens mapping management view | `Views/Pages/BOQManagerPage.xaml` | S |
| TASK-004-11-02 | Create mapping management section or dialog with DataGrid showing all current mappings | `Views/Dialogs/ProductMappingDialog.xaml` | M |
| TASK-004-11-03 | Implement mapping management ViewModel with Add, Delete, and Modify commands | `ViewModels/Dialogs/ProductMappingDialogViewModel.cs` | M |
| TASK-004-11-04 | Bind ProductMappings ObservableCollection to the management view DataGrid | `ViewModels/BOQManagerViewModel.cs` | S |
| TASK-004-11-05 | Implement GetProductMappings() to return all current mappings for display | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-11-06 | Implement Add/Remove mapping operations through IBOQProcessor interface | `OdooAutoCAD.Core/BOQ/IBOQProcessor.cs` | S |
| TASK-004-11-07 | Wire Odoo product search for adding/modifying mappings | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |

## Dependencies
- Depends on: US-004-08 (product mapping infrastructure must exist)
- Blocks: None

## Notes
- The mapping management view can be implemented either as a separate tab/section within the Product Mapping dialog or as a standalone dialog accessible from the BOQ Manager page.
- The underlying data structure is a `Dictionary<string, int>` with `StringComparer.OrdinalIgnoreCase` in the `BOQProcessor` class.
- The `ProductMappings` `ObservableCollection<ProductMappingEntry>` in the ViewModel provides the display-ready data, where each entry has `AutoCADName`, `OdooProductId`, and `OdooProductName`.
- Currently mappings persist only for the application lifecycle (in-memory). Future enhancements could add SQLite persistence for cross-session mapping storage.
- This is a lower-priority (P3) feature since individual product mapping (US-004-08) handles the most common use case. This story provides bulk management for system administrators.
