# US-010-04: Runtime Mode Switching

## User Story
**As a** System Admin,
**I want to** switch between COM and File mode at runtime,
**So that** I can adapt to different workstation configurations.

## Parent Feature
- **FR**: [FR-010-dual-mode-autocad](FR-010-dual-mode-autocad.md)
- **Priority**: P1

## Acceptance Criteria
- [ ] AC-01: Changing mode in Settings immediately updates the DrawingDataServiceDispatcher
- [ ] AC-02: Mode switch disconnects COM / unloads file before switching
- [ ] AC-03: All ViewModels reflect the new mode without requiring page navigation
- [ ] AC-04: Dashboard connection buttons adapt to current mode (Connect vs Open DWG)
- [ ] AC-05: Existing tests pass in both COM and File mode (mocked)
- [ ] AC-06: BOQViewModel uses IDrawingDataService instead of direct IGUIProxy calls
- [ ] AC-07: PurchaseRequisitionViewModel uses IDrawingDataService for header ID extraction
- [ ] AC-08: ParameterConfigViewModel uses IDrawingDataService for attribute read/write
- [ ] AC-09: DashboardViewModel uses IDrawingDataService for connect/load

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-010-003 | Mode change SHALL take effect immediately without restart | Must |
| FR-010-005 | All ViewModels SHALL consume IDrawingDataService | Must |
| FR-010-009 | Runtime dispatcher SHALL route calls based on current mode | Must |
| FR-010-022 | AutoCAD page SHALL show mode indicator | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-010-04-01 | Implement ComDrawingDataService wrapping IAutoCADService + IGUIProxy | `OdooAutoCAD.Core/AutoCAD/ComDrawingDataService.cs` | M |
| TASK-010-04-02 | Refactor BOQViewModel → IDrawingDataService | `ViewModels/BOQViewModel.cs` | M |
| TASK-010-04-03 | Refactor PurchaseRequisitionViewModel → IDrawingDataService | `ViewModels/PurchaseRequisitionViewModel.cs` | M |
| TASK-010-04-04 | Refactor ParameterConfigViewModel → IDrawingDataService | `ViewModels/ParameterConfigViewModel.cs` | M |
| TASK-010-04-05 | Refactor DashboardViewModel → mode-aware connect/load | `ViewModels/DashboardViewModel.cs` | M |
| TASK-010-04-06 | Update BOQViewModelTests for dual-mode (mock IDrawingDataService) | `tests/` | M |
| TASK-010-04-07 | Update PurchaseRequisitionViewModelTests for dual-mode | `tests/` | S |
| TASK-010-04-08 | Update ParameterConfigViewModelTests for dual-mode | `tests/` | S |
| TASK-010-04-09 | Tests: ComDrawingDataService delegation | `tests/` | M |

## Dependencies
- Depends on: US-010-01 (mode infrastructure)
- Depends on: US-010-02 (file read operations)
- Depends on: US-010-03 (file write operations)
- Blocks: None

## Notes
- The ViewModel refactoring is the largest task. Each ViewModel currently calls `_guiProxy.ExecuteInGuiAsync("handler_name", ...)` directly. These calls must be replaced with `_drawingDataService.MethodAsync()` calls.
- `ComDrawingDataService` wraps the existing pattern: it still calls IGUIProxy internally, but ViewModels no longer know about it.
- `FileDrawingDataService` has no COM dependency and can run on any thread.
- The `DrawingDataServiceDispatcher` is a thin proxy that routes to the active backend. It also fires a `ModeChanged` event for UI updates.
- DashboardViewModel needs special attention: the "Connect AutoCAD" button should become "Open DWG..." in file mode.
- Existing test mocks for IGUIProxy need to be replaced with IDrawingDataService mocks, but the test logic remains similar.
