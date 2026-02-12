# US-009-03: Write Parameters to AutoCAD

## User Story
**As a** CAD Engineer,
**I want to** write selected parameter values to AutoCAD attribute blocks,
**So that** the drawing attributes are updated with the correct Odoo data.

## Parent Feature
- **FR**: [FR-009-parameter-config](FR-009-parameter-config.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: Submit command writes all 7 attribute values (product_name, spec, product_catelog, operation_flow, surface_treatment, color_name, color_no) to AutoCAD
- [ ] AC-02: The `autocad_get_attribute_block` handler finds the attribute block reference in a layout
- [ ] AC-03: The `autocad_set_attribute_values` handler writes a dictionary of tag→value pairs to a block
- [ ] AC-04: Attributes are written to all layouts in the current drawing (matching Python behavior)
- [ ] AC-05: All COM operations execute on the GUI/STA thread via IGUIProxy
- [ ] AC-06: Success feedback is displayed: "已成功更新 {count} 個佈局的參數。"
- [ ] AC-07: Failure feedback is displayed with error details
- [ ] AC-08: Submit requires AutoCAD to be connected (disabled otherwise)
- [ ] AC-09: project_id is resolved from the current drawing's layout attributes for color API calls

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-009-015 | Submit SHALL write all 7 attribute values to the attribute block in each layout | Must |
| FR-009-016 | Write operation SHALL execute on the GUI/STA thread via IGUIProxy | Must |
| FR-009-017 | Write operation SHALL require AutoCAD to be connected | Must |
| FR-009-018 | Write operation SHALL display success/failure feedback | Must |
| FR-009-019 | Write operation SHALL find the attribute block via `get_attribute_block` equivalent | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-009-03-01 | Implement `autocad_get_attribute_block` GUIProxy handler | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-009-03-02 | Implement `autocad_set_attribute_values` GUIProxy handler | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | M |
| TASK-009-03-03 | Implement SubmitCommand in ParameterConfigViewModel | `ViewModels/ParameterConfigViewModel.cs` | M |
| TASK-009-03-04 | Implement project_id resolution flow for color API | `ViewModels/ParameterConfigViewModel.cs` | M |
| TASK-009-03-05 | Unit tests for SubmitCommand (success, failure, disabled states) | `tests/.../ParameterConfigViewModelTests.cs` | M |
| TASK-009-03-06 | Unit tests for handler registration (get_attribute_block, set_attribute_values) | `tests/.../AutoCADHandlerTests.cs` | S |

## Dependencies
- Depends on: US-009-02 (form UI must exist with field selections)
- Depends on: Sprint 3 infrastructure (AutoCADService, IGUIProxy, STA threading)
- Blocks: None (terminal user story)

## Notes

### Python Submit Flow
```python
def submit(self):
    # 1. Validate all fields
    if not all([material, spec, category, process, surface, color, color_no]):
        log("請填寫所有欄位。")
        return

    # 2. Build attribute dictionary
    attr_values = {
        'product_name': material,
        'spec': spec,
        'product_catelog': category,
        'operation_flow': process,
        'surface_treatment': surface,
        'color_name': color,
        'color_no': color_no
    }

    # 3. Iterate all layouts
    for layout in layouts:
        block = get_attribute_block(layout)
        if block:
            for tag, val in attr_values.items():
                set_attribute_value(block, tag, val)
```

### C# Implementation Pattern
- The `autocad_get_attribute_block` handler mirrors Python's `get_attribute_block(layout)`: iterates block references in a layout, finds the one with matching attribute tags.
- The `autocad_set_attribute_values` handler accepts a dictionary parameter `{ "layout": "name", "values": { "tag": "value", ... } }` and writes all attributes in one call.
- The SubmitCommand iterates layouts on the C# side, calling the GUIProxy handler for each layout.
- `project_id` is resolved by reading the `pr_no` attribute from the current layout, then looking up the project in Odoo (reusing existing `GetProjectAsync` or extracting from layout attributes).
- All handlers use `Task.FromResult` (sync) since COM attribute operations are fast and must run on STA thread.
