# FR-001: Dashboard Page

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: Not Started
> **Priority**: P1

## 1. Overview

The Dashboard is the default landing page displayed when the application starts. It provides an at-a-glance view of system connection status (AutoCAD, Odoo, MCP Server), quick-action shortcuts for common operations, and recent activity summary. It serves as the central hub from which users orient themselves and navigate to specific features.

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-001-01 | CAD Engineer | See connection status for AutoCAD and Odoo at a glance | I know if I can start working immediately |
| US-001-02 | CAD Engineer | Have quick-connect buttons on the dashboard | I can establish connections without navigating to settings |
| US-001-03 | Project Manager | See recent activity and sync status | I know the current state of data flow between systems |
| US-001-04 | System Admin | See MCP server status | I can verify AI assistant availability |

## 3. Python Reference

### Source Files
- `forms/form_main_modern.py` - Welcome screen (default view when no action selected)
- `forms/form_main_modern.py` - Top bar with connection indicators

### Key Functions
- `create_main_content()` → Creates welcome screen with quick-start buttons
- `update_connection_status()` → Updates Odoo/AutoCAD status labels and button colors
- `on_mcp_sse_status_update(message, is_running)` → SSE status callback

### Python UI Elements
- Welcome title label: "歡迎使用 AutoCAD Odoo 整合系統"
- Welcome instruction text
- Quick Connect Odoo button (150x40px)
- Quick Connect AutoCAD button (150x40px)
- Odoo status label in top bar: "🏢 Odoo: ✅ 已連接" / "🏢 Odoo: 未連接"
- AutoCAD status label in top bar: "📐 AutoCAD: ✅ 已連接" / "📐 AutoCAD: 未連接"
- SSE status indicator: 🟢 (running) / 🔴 (stopped)

## 4. Functional Requirements

### Connection Status Display

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-001-001 | Dashboard SHALL display AutoCAD connection status as a card with icon, label, and connected/disconnected state | Must |
| FR-001-002 | Dashboard SHALL display Odoo connection status as a card with icon, label, and connected/disconnected state | Must |
| FR-001-003 | Dashboard SHALL display MCP Server status as a card with running/stopped state and port info | Should |
| FR-001-004 | Connection status cards SHALL update in real-time when connection state changes | Must |
| FR-001-005 | Connected state SHALL display with green indicator; disconnected with red/gray indicator | Must |

### Quick Actions

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-001-006 | Dashboard SHALL provide a "Connect to AutoCAD" quick-action button | Must |
| FR-001-007 | Dashboard SHALL provide a "Connect to Odoo" quick-action button | Must |
| FR-001-008 | Quick-action buttons SHALL be disabled and show "Connected" state when already connected | Should |
| FR-001-009 | Dashboard SHOULD provide shortcuts to common workflows (Get Parameters, Push to BOQ) | Could |

### Activity Summary

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-001-010 | Dashboard SHOULD display last sync timestamp for Odoo data | Should |
| FR-001-011 | Dashboard SHOULD display count of recent operations (parameters extracted, BOQs pushed) | Could |
| FR-001-012 | Dashboard SHOULD display current project info (PR No, Project Name) when AutoCAD is connected | Should |

## 5. UI Wireframe Description

```
┌──────────────────────────────────────────────────────────┐
│                    Dashboard Page                          │
├──────────────────────────────────────────────────────────┤
│                                                            │
│  Welcome to AutoCAD-Odoo Integration System               │
│  ─────────────────────────────────────────                │
│                                                            │
│  ┌─────────────────┐ ┌─────────────────┐ ┌────────────┐  │
│  │ 📐 AutoCAD      │ │ 🏢 Odoo         │ │ 🤖 MCP     │  │
│  │ ● Connected     │ │ ○ Disconnected  │ │ ● Running  │  │
│  │                 │ │                 │ │ Port: 8084 │  │
│  │ [Connect]       │ │ [Connect]       │ │ [Toggle]   │  │
│  └─────────────────┘ └─────────────────┘ └────────────┘  │
│                                                            │
│  Current Project                                          │
│  ┌──────────────────────────────────────────────────────┐ │
│  │ PR No: PR-2025-001  │  Project: Steel Frame Phase 2 │ │
│  │ Job Plan: JP-001     │  Last Sync: 2025-08-19 14:30 │ │
│  └──────────────────────────────────────────────────────┘ │
│                                                            │
│  Quick Actions                                            │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌────────────┐  │
│  │📋 Get    │ │📊 Push   │ │🔄 Transfer│ │🌊 AI       │  │
│  │Parameters│ │to BOQ    │ │BOQ to PR │ │Assistant   │  │
│  └──────────┘ └──────────┘ └──────────┘ └────────────┘  │
│                                                            │
└──────────────────────────────────────────────────────────┘
```

### Layout Details
- **Connection Cards**: 3 cards in a horizontal row using WPF `UniformGrid` or `StackPanel`
- **Each Card**: Border with rounded corners, icon, title, status indicator (green dot/red dot), action button
- **Project Info**: Panel showing current project context from AutoCAD
- **Quick Actions**: Row of shortcut buttons navigating to other pages or triggering actions

## 6. Data Model

### Input Data (ViewModel Properties)

```csharp
public class DashboardViewModel : ObservableObject
{
    // Connection Status
    public bool IsAutoCADConnected { get; set; }
    public string AutoCADStatus { get; set; }     // "Connected" / "Disconnected"
    public string AutoCADDocumentName { get; set; }

    public bool IsOdooConnected { get; set; }
    public string OdooStatus { get; set; }
    public string OdooServerUrl { get; set; }

    public bool IsMCPRunning { get; set; }
    public string MCPStatus { get; set; }
    public int MCPPort { get; set; }

    // Project Info
    public string CurrentPRNo { get; set; }
    public string CurrentProjectName { get; set; }
    public string CurrentJobPlanName { get; set; }
    public DateTime? LastSyncTime { get; set; }

    // Commands
    public IRelayCommand ConnectAutoCADCommand { get; }
    public IRelayCommand ConnectOdooCommand { get; }
    public IRelayCommand ToggleMCPCommand { get; }
    public IRelayCommand NavigateToParametersCommand { get; }
    public IRelayCommand NavigateToBOQCommand { get; }
    public IRelayCommand NavigateToPRCommand { get; }
}
```

## 7. API/Service Dependencies

| Service | Interface | Usage |
|---------|-----------|-------|
| AutoCAD Service | `IAutoCADService` | `IsConnected`, `ConnectAsync()`, `GetStatusAsync()` |
| Odoo Service | `IOdooService` | `IsConnected`, `ConnectAsync()`, `GetStatusAsync()`, `GetLastSyncTime()` |
| Navigation Service | `INavigationService` | `NavigateTo()` for quick-action buttons |
| MCP SSE Manager | (MCP project) | `IsRunning`, `StartServer()`, `StopServer()`, `GetServerStatus()` |

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-001-001 | Quick-connect buttons must check current connection state before attempting connection |
| VR-001-002 | Quick-action workflow buttons (Get Parameters, Push BOQ) should only be enabled when required connections are active |
| VR-001-003 | "Get Parameters" requires both AutoCAD AND Odoo connected |
| VR-001-004 | "Push to BOQ" requires Odoo connected and project context available |

## 9. Error Handling

| Scenario | User-Facing Message | Action |
|----------|---------------------|--------|
| AutoCAD connect fails | "Unable to connect to AutoCAD. Please ensure AutoCAD is running." | Show in status card, enable retry |
| Odoo connect fails | "Unable to connect to Odoo server. Check connection settings." | Show in status card, enable retry |
| MCP server start fails | "MCP Server failed to start on port {port}. Port may be in use." | Show error in MCP card |
| No project found | "No project context available. Open a drawing with PR information." | Show in project info panel |

## 10. Implementation Notes

### C# Target Files
- `Views/Pages/DashboardPage.xaml` - WPF page with data bindings
- `ViewModels/DashboardViewModel.cs` - MVVM ViewModel with commands and properties

### MVVM Bindings
- Use `CommunityToolkit.Mvvm` for `ObservableObject`, `RelayCommand`
- Connection status bound via `{Binding IsAutoCADConnected}` with `BoolToVisibility` converters
- Status indicator colors via `{Binding IsAutoCADConnected, Converter={StaticResource BoolToColorConverter}}`
- Command bindings: `{Binding ConnectAutoCADCommand}` on button Click

### Special Considerations
- Connection status should use event subscriptions to `IAutoCADService` and `IOdooService` status change events
- Dashboard ViewModel should be registered as singleton in DI container to maintain state
- Consider using `DispatcherTimer` (100ms) for polling GUI proxy requests (matching Python pattern)
- MCP status requires async monitoring of SSE server state

### Differences from Python
- Python uses emoji indicators (🟢/🔴) → C# uses Ellipse with Fill color binding
- Python welcome screen is simple → C# dashboard is richer with cards pattern
- Python status is in top bar → C# dashboard consolidates status into dedicated page
