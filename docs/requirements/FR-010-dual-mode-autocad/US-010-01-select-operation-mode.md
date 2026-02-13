# US-010-01: Select Operation Mode

## User Story
**As a** CAD Engineer,
**I want to** select between COM and File operation mode,
**So that** I can use the application with AutoCAD LT or full AutoCAD.

## Parent Feature
- **FR**: [FR-010-dual-mode-autocad](FR-010-dual-mode-autocad.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Settings > AutoCAD tab displays two RadioButtons: "COM 模式" and "檔案模式"
- [ ] AC-02: COM mode is selected by default on fresh install
- [ ] AC-03: Selected mode persists across application restarts via ISettingsService
- [ ] AC-04: Mode change triggers `DrawingDataServiceDispatcher` to switch backend
- [ ] AC-05: SettingsViewModel exposes `AutoCADOperationMode` property with two-way binding
- [ ] AC-06: AutoCAD page shows current mode indicator ("Mode: COM" or "Mode: File")
- [ ] AC-07: Mode descriptions explain the difference (COM = live connection, File = open DWG)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-010-001 | Settings SHALL provide COM / File mode selection | Must |
| FR-010-002 | Selected mode SHALL persist across application restarts | Must |
| FR-010-004 | COM mode SHALL remain the default for backwards compatibility | Must |
| FR-010-025 | Settings > AutoCAD tab SHALL contain mode selection RadioButtons | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-010-01-01 | Define AutoCADOperationMode enum and WriteStrategy enum | `OdooAutoCAD.Core/AutoCAD/IDrawingDataService.cs` | S |
| TASK-010-01-02 | Define IDrawingDataService interface with 12 methods | `OdooAutoCAD.Core/AutoCAD/IDrawingDataService.cs` | M |
| TASK-010-01-03 | Add AutoCADOperationMode property to SettingsViewModel + ISettingsService | `ViewModels/SettingsViewModel.cs` | S |
| TASK-010-01-04 | Add mode selection RadioButtons to Settings > AutoCAD tab | `Views/Pages/SettingsPage.xaml` | S |
| TASK-010-01-05 | Implement DrawingDataServiceDispatcher with mode switching | `OdooAutoCAD.Core/AutoCAD/DrawingDataServiceDispatcher.cs` | M |
| TASK-010-01-06 | Register new DI services in App.xaml.cs | `App.xaml.cs` | S |
| TASK-010-01-07 | Add mode indicator to AutoCAD page | `Views/Pages/AutoCADPage.xaml` | S |
| TASK-010-01-08 | Tests: Dispatcher mode switching and Settings persistence | `tests/` | M |

## Dependencies
- Depends on: Sprint 2 (SettingsViewModel, ISettingsService infrastructure)
- Depends on: Sprint 3 (IAutoCADService, IGUIProxy infrastructure)
- Blocks: US-010-02, US-010-03, US-010-04 (mode infrastructure required)

## Notes
- COM mode is the established default. File mode is additive — COM behavior must remain unchanged.
- The `DrawingDataServiceDispatcher` holds references to both backends and delegates based on current mode.
- Mode switch clears the current connection/loaded file state to avoid stale data.
- The dispatcher is registered as Singleton in DI; backends are Transient.
