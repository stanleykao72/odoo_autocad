# US-009-02: Parameter Configuration Form UI

## User Story
**As a** CAD Engineer,
**I want to** see a parameter configuration form with searchable dropdowns,
**So that** I can select the correct material, processing, and color values.

## Parent Feature
- **FR**: [FR-009-parameter-config](FR-009-parameter-config.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: A "參數配置" page is accessible via a sidebar navigation button
- [ ] AC-02: The page displays 3 GroupBox sections: 材料配置, 加工配置, 顏色配置
- [ ] AC-03: Material field is a searchable ComboBox populated from product API data
- [ ] AC-04: Unit field is read-only and auto-fills when a material is selected
- [ ] AC-05: Spec, Category, Operation Flow, Surface Treatment are searchable ComboBoxes
- [ ] AC-06: Color field is a searchable ComboBox; Color No is read-only and auto-fills on color selection
- [ ] AC-07: All dropdown options are loaded in parallel on page load via LoadOptionsCommand
- [ ] AC-08: A loading indicator is shown while options are being fetched
- [ ] AC-09: Submit button is enabled only when all 7 required fields are filled
- [ ] AC-10: Cancel button clears all selections and resets the form
- [ ] AC-11: Status message area displays loading/success/error feedback

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-009-007 | Page SHALL display a 3-section form: 材料配置, 加工配置, 顏色配置 | Must |
| FR-009-008 | Material field SHALL be a searchable ComboBox | Must |
| FR-009-009 | Unit field SHALL be read-only, auto-filled on material selection | Must |
| FR-009-010 | Spec, Category, Op Flow, Surface Treatment SHALL be searchable ComboBoxes | Must |
| FR-009-011 | Color field SHALL be searchable; Color No SHALL be read-only auto-filled | Must |
| FR-009-012 | Submit button SHALL be enabled only when all 7 required fields are filled | Must |
| FR-009-013 | Cancel button SHALL clear all fields and reset the form | Should |
| FR-009-014 | Page SHALL be accessible via sidebar navigation button | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-009-02-01 | Create ParameterConfigViewModel with all properties and collections | `ViewModels/ParameterConfigViewModel.cs` | L |
| TASK-009-02-02 | Create ParameterConfigPage.xaml with 3 GroupBox sections | `Views/Pages/ParameterConfigPage.xaml` | L |
| TASK-009-02-03 | Code-behind with DI resolution pattern | `Views/Pages/ParameterConfigPage.xaml.cs` | S |
| TASK-009-02-04 | Add sidebar navigation button for 參數配置 | `Views/MainWindow.xaml` | S |
| TASK-009-02-05 | DI registration + page title mapping for ParameterConfig | `App.xaml.cs`, `ViewModels/MainViewModel.cs` | S |
| TASK-009-02-06 | Implement parallel API loading in LoadOptionsCommand | `ViewModels/ParameterConfigViewModel.cs` | M |
| TASK-009-02-07 | Implement selection changed auto-fill handlers (Unit, ColorNo) | `ViewModels/ParameterConfigViewModel.cs` | S |
| TASK-009-02-08 | Implement AllFieldsFilled computed property + CanExecute for Submit | `ViewModels/ParameterConfigViewModel.cs` | S |
| TASK-009-02-09 | ViewModel unit tests for form state, loading, selection, validation | `tests/.../ParameterConfigViewModelTests.cs` | M |

## Dependencies
- Depends on: US-009-01 (API methods must exist to populate dropdowns)
- Depends on: Sprint 1 infrastructure (Navigation, DI, sidebar pattern)
- Blocks: US-009-03 (Submit command writes to AutoCAD)

## Notes
- The Python form uses `PopupSelector` (a custom Tkinter popup with search). The C# equivalent is `ComboBox` with `IsEditable="True"` and `IsTextSearchEnabled="True"` for type-ahead filtering.
- Python fetches data lazily on field focus events. C# loads all options in parallel on page load for a smoother UX.
- The page follows the established pattern: code-behind resolves ViewModel via `App.Services.GetRequiredService<ParameterConfigViewModel>()`, sets `DataContext`, and calls `LoadOptionsCommand.Execute(null)`.
- Material selection auto-fills Unit from the `OdooProduct.Unit` property (read-only TextBox).
- Color selection auto-fills Color No from the `OdooColor.ColorNo` property (read-only TextBox).
- `AllFieldsFilled` is a computed bool: all 5 ComboBox selections are non-null AND both read-only fields are non-empty.
