# US-002-01: Connect to AutoCAD

## User Story
**As a** CAD Engineer,
**I want to** connect to a running AutoCAD instance,
**So that** I can extract parameters from my current drawing.

## Parent Feature
- **FR**: [FR-002-autocad-connection](FR-002-autocad-connection.md)
- **Priority**: P1

## Acceptance Criteria
- [x] AC-01: Clicking the Connect button initiates a COM connection to AutoCAD
- [x] AC-02: Connection first attempts GetActiveObject to attach to a running instance; if that fails, falls back to Dispatch to launch a new instance
- [x] AC-03: Connection retries up to 5 times (1-second intervals) to obtain ActiveDocument
- [x] AC-04: AutoCAD application is set to Visible upon successful connection
- [x] AC-05: Connection status indicator updates to "Connected" with a visual indicator (e.g., green icon) on success
- [x] AC-06: Connect button changes to "Disconnect" after a successful connection
- [x] AC-07: All COM operations execute on the GUI/STA thread via IGUIProxy
- [x] AC-08: If AutoCAD is not running and cannot be launched, an error message is displayed: "AutoCAD is not running. Please start AutoCAD and try again."
- [x] AC-09: If connection fails after retries, an error message is displayed: "Could not access the active document after 5 attempts."

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-002-001 | Page SHALL provide a Connect button that connects to AutoCAD via COM | Must |
| FR-002-002 | Connection SHALL first try GetActiveObject, then fallback to Dispatch | Must |
| FR-002-003 | Page SHALL display connection status (Connected/Disconnected) with visual indicator | Must |
| FR-002-004 | Page SHALL display AutoCAD version and application info when connected | Should |
| FR-002-005 | Page SHALL display current document filename and full path | Must |
| FR-002-006 | Connect button SHALL change appearance when connected (e.g., show Disconnect) | Should |
| FR-002-007 | All COM operations SHALL execute on the GUI/STA thread via IGUIProxy | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-002-01-01 | Create AutoCAD page XAML with connection panel (Connect/Disconnect button, status indicator) | `Views/Pages/AutoCADPage.xaml` | M |
| TASK-002-01-02 | Implement ConnectCommand and DisconnectCommand in ViewModel | `ViewModels/AutoCADViewModel.cs` | M |
| TASK-002-01-03 | Implement ConnectAsync in AutoCADService with GetActiveObject/Dispatch fallback and retry logic | `OdooAutoCAD.Core/AutoCAD/AutoCADService.cs` | L |
| TASK-002-01-04 | Define ConnectAsync and DisconnectAsync in IAutoCADService interface | `OdooAutoCAD.Core/AutoCAD/IAutoCADService.cs` | S |
| TASK-002-01-05 | Wire COM operations through IGUIProxy.ExecuteInGuiAsync for STA thread safety | `OdooAutoCAD.Core/Threading/IGUIProxy.cs` | M |
| TASK-002-01-06 | Implement connection status properties (IsConnected, Status, DocumentName, DocumentPath) in ViewModel | `ViewModels/AutoCADViewModel.cs` | S |

## Dependencies
- Depends on: US-008-01 (navigation framework)
- Blocks: US-002-02, US-002-03, US-002-06

## Notes
- The COM connection pattern in C# uses `Marshal.GetActiveObject("AutoCAD.Application")` as the primary method and `Activator.CreateInstance(Type.GetTypeFromProgID("AutoCAD.Application"))` as fallback, mirroring the Python `GetActiveObject`/`Dispatch` pattern.
- All COM calls must be marshalled to the STA thread. WPF uses a `DispatcherTimer` (100ms interval) to call `IGUIProxy.ProcessRequests()`.
- Python uses `pythoncom.CoInitialize()` for COM threading; in C# this is handled by the STA thread model inherent to WPF's Dispatcher.
- Validation rule VR-002-001: Connect button only enabled when not connected. VR-002-002: All COM-dependent controls disabled when disconnected.

### COM Threading Fix (2026-02-11) — Pure STA Mode

A critical COM threading issue was identified and fixed. The original implementation used `Task.Run()` inside GUIProxy handlers, pushing COM operations to MTA thread pool threads while blocking the STA thread with `.GetAwaiter().GetResult()`. This caused cross-apartment marshaling deadlocks and Access Violations when connecting to AutoCAD.

**Root causes fixed**:
1. **Removed `Task.Run` from all 10 GUIProxy handlers** — COM operations now execute directly on the WPF STA thread. The IDispatch QI warmup (`GetIDispatchForObject` after `GetActiveObject`) prevents deferred cross-process QI crashes.
2. **Registered `OleMessageFilter`** on STA thread at startup (`App.xaml.cs`) — automatically retries COM calls rejected by AutoCAD during WPF layout processing (`RPC_E_CALL_REJECTED` / `0x80010001`).
3. **GUIProxy async-aware `ProcessSingleRequest`** — uses fast-path (sync handlers via `Task.FromResult`) and slow-path (`ContinueWith` on STA `SynchronizationContext`) to avoid blocking the STA thread on async handlers.
4. **`ConnectInternalAsync`** uses `await Task.Delay(1000)` instead of `Thread.Sleep(1000)` for retries, keeping the UI responsive during the 5-attempt retry loop.
5. **Safety improvements**: `volatile bool _isConnected`, `Marshal.IsComObject()` guard before `ReleaseComObject()`.

**Files modified**: `AutoCADService.cs`, `GUIProxy.cs`, `OleMessageFilter.cs` (internal→public), `App.xaml.cs`
