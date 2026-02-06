# FR-008: UI Framework & Navigation

> **Document Version**: 1.0
> **Last Updated**: 2026-02-06
> **Status**: In Progress
> **Priority**: P0

## 1. Overview

The UI Framework & Navigation system defines the WPF application shell: the `MainWindow` that hosts all page content through Frame-based navigation, the sidebar navigation panel, the header bar with context-sensitive page titles, connection status indicators, a system log panel, and theming infrastructure. This is the top-level architectural container within which all feature pages (Dashboard, AutoCAD, Odoo, BOQ, PR, MCP, Settings) are rendered.

In the Python implementation, `ModernFormMain(ctk.CTk)` serves as the single-window shell using CustomTkinter. It uses a grid layout with a top bar (row 0), sidebar + main content (row 1), and bottom log area (row 2). Navigation is achieved by destroying and recreating widgets in the main content area. The C# port replaces this with a proper WPF `Frame`-based navigation model where `MainWindow.xaml` contains a persistent sidebar, header, status bar, and a `Frame` element (`MainFrame`) that loads distinct `Page` objects via `NavigationService`.

The framework also encompasses:
- Application lifecycle management (startup, shutdown, cleanup)
- Dependency injection container configuration (`App.xaml.cs`)
- GUI Proxy timer for thread-safe COM operations (100ms `DispatcherTimer`)
- Resource dictionaries for theming and styling (`App.xaml`)
- CJK font rendering support (Microsoft JhengHei UI)
- Global exception handling

## 2. User Stories

| ID | As a... | I want to... | So that... |
|----|---------|-------------|------------|
| US-008-01 | CAD Engineer | Navigate between different functional areas using a sidebar menu | I can quickly access AutoCAD, Odoo, BOQ, and PR features without losing context |
| US-008-02 | CAD Engineer | See which page I am currently on via visual highlighting in the sidebar | I always know my current location in the application |
| US-008-03 | CAD Engineer | See connection status for AutoCAD, Odoo, and MCP at all times regardless of which page I am on | I am always aware of system connectivity |
| US-008-04 | CAD Engineer | View system log messages at the bottom of the screen | I can monitor operations and diagnose issues in real-time |
| US-008-05 | System Admin | Have the application start centered on screen at a reasonable default size | The window is immediately usable without manual resizing |
| US-008-06 | System Admin | Have the MCP server and GUI proxy start automatically when the application launches | AI assistant integration is available without manual intervention |
| US-008-07 | CAD Engineer | Use keyboard shortcuts to navigate between pages | I can work efficiently without relying solely on mouse interaction |
| US-008-08 | CAD Engineer | See the current page title in the header area | I have clear confirmation of what page is displayed |
| US-008-09 | System Admin | Have all resources cleaned up properly when I close the application | No orphan processes or port locks are left behind |
| US-008-10 | CAD Engineer | Have CJK characters render correctly throughout the application | All Chinese labels, status messages, and log entries are legible |

## 3. Python Reference

### Source Files
- `forms/form_main_modern.py` - `ModernFormMain(ctk.CTk)` main window class, overall shell layout
- `ui/ui_theme.py` - `UITheme` class with COLORS, FONTS, SIZES, ICONS dictionaries
- `ui/ui_fonts.py` - `FontManager` class with cross-platform font fallback system
- `utility/util_gui_proxy.py` - `GUIProxy` class with message queue for thread-safe COM operations

### Layout Structure (Grid)
```
Row 0: Top Bar          (col 0..1)  - Title + connection status indicators + SSE controls
Row 1: Sidebar (col 0)  + Main Content (col 1, weight=1)
Row 2: Bottom Area       (col 0..1)  - System log panel
```

- `grid_columnconfigure(1, weight=1)` - Main content column expands
- `grid_rowconfigure(1, weight=1)` - Content row expands

### Navigation Pattern (Python)
Python does not use page-based navigation. Instead, `clear_main_content()` destroys all children of `self.main_content` and recreates the welcome screen. Feature-specific content is rendered by replacing widgets in the main content frame:

```python
def clear_main_content(self):
    for widget in self.main_content.winfo_children():
        widget.destroy()
    self.create_welcome_screen()
```

Sidebar buttons call action methods directly (e.g., `connect_odoo`, `get_parameters_from_odoo`, `push_to_boq`) which either execute logic or replace main content area widgets.

### Initialization Sequence (Python)
1. `__init__()` - Set theme, store connection config, set title and icon
2. `setup_window_geometry()` - Min 1000x700, default 80% of screen (max 1200x800), centered
3. `setup_window_style()` - WM_DELETE_WINDOW protocol, deferred icon load
4. `create_ui()` - Build top bar, sidebar, main content, bottom area
5. `setup_log_util()` - Initialize UtilLog in bottom frame
6. `init_utilities()` - Create OdooUtil, AutoCADUtil, BOQ/PR utils, MCPSSEManager, GUI Proxy
7. `start_gui_proxy_processing()` - Begin 100ms timer for proxy request processing
8. `auto_start_sse_server()` - Launch SSE server in background thread
9. `update_connection_status()` - Initial status display

### Key Python UI Details
- **Title**: "AutoCAD Odoo 整合系統"
- **Window icon**: `icon/odoo_autocad.ico`
- **Min size**: 1000x700, default 80% screen (max 1200x800), centered
- **Top bar title**: "AutoCAD Odoo 整合系統" with primary color background, white text
- **Sidebar title**: "功能選單" (18pt bold)
- **Sidebar sections**: "連接管理" (Connect), "主要功能" (Main), "工具" (Tools), SSE Control
- **Button colors**: Odoo blue (#1976D2), AutoCAD purple (#7B1FA2), green (#388E3C), orange (#F57C00), brown (#5D4037), warning (#FF8F00), danger (#D32F2F), connected (#2E7D32), SSE (#FF9800)
- **Status indicators**: SSE 🟢/🔴, Odoo/AutoCAD text labels with emoji
- **Bottom panel**: "系統日誌" heading, ScrolledText (black bg, white text)
- **Window close**: `on_closing()` stops MCP/SSE servers, destroys window
- **GUI proxy polling**: `self.after(100, self.process_gui_proxy_requests)` every 100ms

## 4. Functional Requirements

### Main Window Layout

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-001 | MainWindow SHALL use a two-column Grid layout: fixed-width sidebar (left, 250px) and flexible content area (right, `*`) | Must |
| FR-008-002 | The content area SHALL contain three rows: header (Auto), content Frame (`*`), and status bar (Auto) | Must |
| FR-008-003 | MainWindow SHALL set a default size of 1280x720 with a minimum size of 900x600 | Must |
| FR-008-004 | MainWindow SHALL use `WindowStartupLocation="CenterScreen"` to center the window on launch | Must |
| FR-008-005 | MainWindow SHALL set its background to `BackgroundBrush` from the application resource dictionary | Must |

### Side Navigation

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-006 | The sidebar SHALL display a title area at the top with the application name "Odoo AutoCAD" and version "Integration System v6.0" | Must |
| FR-008-007 | The sidebar SHALL contain navigation buttons for: Dashboard, AutoCAD, Odoo, BOQ Manager, Purchase Requisition, AI Assistant, Settings | Must |
| FR-008-008 | Each navigation button SHALL display an icon (emoji or icon font) and a text label, left-aligned | Must |
| FR-008-009 | Navigation buttons SHALL use a consistent `NavButton` style with transparent background, left-aligned content, and left-border accent on hover | Must |
| FR-008-010 | The currently active navigation button SHALL be visually highlighted with a distinct background color and a left border accent in the primary color | Must |
| FR-008-011 | The sidebar SHALL display a Connection Status section at the bottom with AutoCAD, Odoo, and MCP Server status indicators | Must |

### Page Navigation

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-012 | MainWindow SHALL contain a `Frame` element (`MainFrame`) with `NavigationUIVisibility="Hidden"` for hosting page content | Must |
| FR-008-013 | `INavigationService.NavigateTo(pageName)` SHALL resolve pages by reflection from namespace `OdooAutoCAD.App.Views.Pages.{pageName}Page` | Must |
| FR-008-014 | `INavigationService.GoBack()` SHALL navigate to the previous page in the Frame journal when `CanGoBack` is true | Should |
| FR-008-015 | Navigation SHALL default to the Dashboard page on application startup | Must |
| FR-008-016 | Navigation SHALL update `MainViewModel.CurrentPageTitle` to reflect the display name of the active page | Must |
| FR-008-017 | Navigation SHALL update the sidebar button highlighting to match the currently displayed page | Must |

### Top Header Bar

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-018 | The header bar SHALL display `CurrentPageTitle` as a 20pt SemiBold TextBlock bound to the ViewModel | Must |
| FR-008-019 | The header bar SHALL include a "Refresh" button bound to `RefreshCommand` on the right side | Should |
| FR-008-020 | The header bar SHALL include a "Help" button on the right side | Could |
| FR-008-021 | The header bar SHALL use `SurfaceBrush` background with a bottom border | Must |

### Connection Status Indicators

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-022 | The sidebar connection panel SHALL display AutoCAD status with a colored `Ellipse` indicator (green=connected, gray=disconnected) and status text | Must |
| FR-008-023 | The sidebar connection panel SHALL display Odoo status with a colored `Ellipse` indicator (green=connected, gray=disconnected) and status text | Must |
| FR-008-024 | The sidebar connection panel SHALL display MCP Server status with a colored `Ellipse` indicator (green=running with port, gray=stopped) and status text | Must |
| FR-008-025 | Connection indicators SHALL update in real-time via data binding to `MainViewModel` observable properties (`AutoCADStatusColor`, `OdooStatusColor`, `MCPStatusColor`) | Must |
| FR-008-026 | MCP status SHALL display the port number (e.g., "Port 8084") when the server is running | Should |

### System Log Panel

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-027 | The status bar at the bottom of the content area SHALL display a `StatusMessage` text, the current date/time, and the application version | Must |
| FR-008-028 | A dedicated system log panel SHALL be available (either as a collapsible bottom panel or a separate page) to display operational log messages | Should |
| FR-008-029 | The system log panel SHALL auto-scroll to the latest entry | Should |
| FR-008-030 | Log entries SHALL include timestamps and category prefixes (e.g., "[SSE GUI]", "[GUI Proxy]") | Should |

### Theme and Styling

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-031 | App.xaml SHALL define a centralized `ResourceDictionary` containing all color, brush, and style resources | Must |
| FR-008-032 | The color palette SHALL include: PrimaryColor (#1A73E8), PrimaryDarkColor (#1557B0), AccentColor (#00C853), BackgroundColor (#F5F5F5), SurfaceColor (#FFFFFF), ErrorColor (#D32F2F), WarningColor (#FFA000), SuccessColor (#388E3C), TextPrimaryColor (#212121), TextSecondaryColor (#757575), BorderColor (#E0E0E0) | Must |
| FR-008-033 | App.xaml SHALL define named styles: `PrimaryButton` (filled primary, white text, rounded corners), `NavButton` (transparent, left-aligned, hover accent), `CardPanel` (white surface, rounded corners, drop shadow), `ModernTextBox` (bordered, rounded corners, focus highlight) | Must |
| FR-008-034 | All UI elements SHALL use resource references (`{StaticResource BrushName}`) rather than hardcoded colors | Must |
| FR-008-035 | The application SHALL support CJK font rendering using Microsoft JhengHei UI as the primary font family with appropriate fallbacks | Must |

### Responsive Layout

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-036 | The sidebar width SHALL remain fixed at 250px while the content area fills remaining space | Must |
| FR-008-037 | The content Frame SHALL expand to fill all available vertical space between the header and status bar | Must |
| FR-008-038 | The window SHALL be resizable with proper content reflow and minimum size constraints enforced | Must |

### Window Lifecycle

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-039 | On startup, `App.OnStartup` SHALL configure Serilog, build the DI host, initialize the GUI proxy timer, optionally start the MCP server (`--enable-mcp`), and show MainWindow | Must |
| FR-008-040 | On exit, `App.OnExit` SHALL stop the GUI proxy timer, stop the MCP server if running, stop the GUI proxy, stop the DI host, and flush Serilog | Must |
| FR-008-041 | The application SHALL register global exception handlers for `DispatcherUnhandledException`, `AppDomain.UnhandledException`, and `TaskScheduler.UnobservedTaskException` | Must |

### Keyboard Shortcuts

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-008-042 | The application SHOULD support `Ctrl+1` through `Ctrl+7` keyboard shortcuts for navigating to the seven sidebar pages in order | Should |
| FR-008-043 | The application SHOULD support `F5` as a shortcut for the Refresh command | Should |
| FR-008-044 | The application SHOULD support `Alt+Left` for navigating back when `CanGoBack` is true | Could |

## 5. UI Wireframe Description

```
+-----------------------------+------------------------------------------------------+
| SIDEBAR (250px)             | CONTENT AREA (remaining width)                       |
|                             |                                                      |
| +--------------------------+| +---------------------------------------------------+|
| | Odoo AutoCAD             || | HEADER BAR                                        ||
| | Integration System v6.0  || | [CurrentPageTitle]              [Refresh] [Help]   ||
| +--------------------------+| +---------------------------------------------------+|
|                             |                                                      |
| [Dashboard        ]        | +---------------------------------------------------+|
| [AutoCAD          ]        | |                                                   ||
| [Odoo             ]        | |              MAIN FRAME                           ||
| [BOQ Manager      ]        | |                                                   ||
| [Purchase Req.    ]        | |     (Page content loaded via NavigationService)    ||
| [AI Assistant     ]        | |                                                   ||
| [Settings         ]        | |                                                   ||
|                             | |                                                   ||
|                             | +---------------------------------------------------+|
| +--------------------------+|                                                      |
| | Connection Status        || +---------------------------------------------------+|
| | * AutoCAD  Disconnected  || | STATUS BAR                                        ||
| | * Odoo     Disconnected  || | [StatusMessage]        [DateTime]       [v6.0.0]  ||
| | * MCP      Stopped       || +---------------------------------------------------+|
| +--------------------------+|                                                      |
+-----------------------------+------------------------------------------------------+

Optional: Collapsible Log Panel (below status bar or as toggle overlay)
+------------------------------------------------------------------------+
| SYSTEM LOG PANEL                                                       |
| +--------------------------------------------------------------------+ |
| | [2026-02-06 14:30:01] [SSE GUI] SSE server started on port 8084   | |
| | [2026-02-06 14:30:02] [GUI Proxy] GUI proxy initialized           | |
| | [2026-02-06 14:30:05] [Odoo] Connected to server                  | |
| +--------------------------------------------------------------------+ |
+------------------------------------------------------------------------+
```

### Layout Details
- **Sidebar**: Fixed 250px width, `SurfaceBrush` background, right border separator. Three vertical sections: title/logo (top), navigation buttons (middle, scrollable), connection status (bottom).
- **Header Bar**: `SurfaceBrush` background, bottom border, 20pt page title on left, action buttons on right.
- **Main Frame**: `Frame` element with 20px margin, hidden navigation UI chrome. Pages load as `Page` instances.
- **Status Bar**: `SurfaceBrush` background, top border. Three columns: status message (left, fills), date/time (center-right), version (far right).
- **Navigation Buttons**: Full-width, left-aligned text, transparent background, 3px left border accent on hover/active, 20px horizontal + 12px vertical padding.
- **Connection Status**: Bottom panel with labeled colored dots (Ellipse), status text, separated by top border.

## 6. Data Model

### MainViewModel (Primary Shell ViewModel)

```csharp
public partial class MainViewModel : ObservableObject
{
    // Navigation State
    [ObservableProperty] private string _currentPageTitle = "Dashboard";
    [ObservableProperty] private string _activeNavButton = "BtnDashboard";

    // Status Bar
    [ObservableProperty] private string _statusMessage = "Ready";
    [ObservableProperty] private string _currentDateTime;

    // AutoCAD Connection Status
    [ObservableProperty] private string _autoCADStatusText = "Disconnected";
    [ObservableProperty] private Brush _autoCADStatusColor = Brushes.Gray;

    // Odoo Connection Status
    [ObservableProperty] private string _odooStatusText = "Disconnected";
    [ObservableProperty] private Brush _odooStatusColor = Brushes.Gray;

    // MCP Server Status
    [ObservableProperty] private string _mcpStatusText = "Stopped";
    [ObservableProperty] private Brush _mcpStatusColor = Brushes.Gray;

    // Commands
    public IRelayCommand RefreshCommand { get; }
    public IRelayCommand<string> NavigateCommand { get; }
    public IAsyncRelayCommand StartMCPServerCommand { get; }
    public IAsyncRelayCommand StopMCPServerCommand { get; }
}
```

### ConnectionStatusViewModel (Embedded in Sidebar)

```csharp
public partial class ConnectionStatusViewModel : ObservableObject
{
    [ObservableProperty] private bool _isAutoCADConnected;
    [ObservableProperty] private bool _isOdooConnected;
    [ObservableProperty] private bool _isMCPRunning;
}
```

### Page Title Mapping

| NavButton Name | Page Name | Display Title |
|----------------|-----------|---------------|
| BtnDashboard | Dashboard | Dashboard |
| BtnAutoCAD | AutoCAD | AutoCAD Integration |
| BtnOdoo | Odoo | Odoo Connection |
| BtnBOQ | BOQ | BOQ Manager |
| BtnPR | PR | Purchase Requisition |
| BtnMCP | MCP | AI Assistant (MCP) |
| BtnSettings | Settings | Settings |

## 7. API/Service Dependencies

| Service | Interface | Registration | Usage in Shell |
|---------|-----------|--------------|----------------|
| Navigation Service | `INavigationService` | Singleton | `NavigateTo(pageName)`, `GoBack()`, `CanGoBack`, `Frame` property set by MainWindow |
| GUI Proxy | `IGUIProxy` | Singleton | `Start()`, `Stop()`, `ProcessRequests()` called by DispatcherTimer at 100ms |
| MCP SSE Server | `MCPSSEServer` | Singleton | `IsRunning`, `Port`, `StartAsync()`, `StopAsync()` for status display and lifecycle |
| MCP Tool Registry | `MCPToolRegistry` | Singleton | Constructed with IGUIProxy, optional IAutoCADService, IOdooService, IBOQProcessor |
| AutoCAD Service | `IAutoCADService` | Singleton (optional) | Connection status polling for sidebar indicators |
| Odoo Service | `IOdooService` | Singleton (optional) | Connection status polling for sidebar indicators |
| BOQ Processor | `IBOQProcessor` | Singleton (optional) | Not directly used by shell, but registered for DI availability |
| Logger | `ILogger<T>` | Transient (Serilog) | Structured logging throughout shell lifecycle |
| MainViewModel | `MainViewModel` | Transient | DataContext for MainWindow, manages all shell-level state |

### DI Registration Order (App.xaml.cs ConfigureServices)
1. Core threading: `IGUIProxy` -> `GUIProxy`
2. MCP services: `MCPToolRegistry`, `MCPSSEServer`
3. ViewModels: `MainViewModel`, `ConnectionStatusViewModel`, `BOQViewModel`, `SettingsViewModel`
4. Application services: `INavigationService` -> `NavigationService`

## 8. Validation Rules

| Rule | Description |
|------|-------------|
| VR-008-001 | `NavigateTo(pageName)` must validate that the resolved page type exists before navigating; if the type is null, navigation must be silently skipped or a fallback "page not found" view displayed |
| VR-008-002 | `GoBack()` must check `CanGoBack` before invoking `Frame.GoBack()` to avoid InvalidOperationException |
| VR-008-003 | The `INavigationService.Frame` property must be set before any navigation calls; MainWindow constructor must assign `MainFrame` to the service |
| VR-008-004 | Active navigation button highlighting must be updated atomically with the Frame navigation to prevent desynchronization |
| VR-008-005 | The GUI proxy timer must not be started until the DI host has fully started and `IGUIProxy` is resolved |
| VR-008-006 | Connection status colors must only transition between defined states: Gray (disconnected/stopped), Green (connected/running), Red (error) |
| VR-008-007 | The `--enable-mcp` flag must be checked case-insensitively against `e.Args` |
| VR-008-008 | Window minimum size (900x600) must be enforced by `MinHeight`/`MinWidth` properties and must not be overridable at runtime |

## 9. Error Handling

| Scenario | User-Facing Behavior | Technical Action |
|----------|---------------------|------------------|
| DI host build fails | MessageBox with startup error message and stack trace; application exits with code 1 | `Log.Fatal()`, `Shutdown(1)` |
| Unhandled dispatcher exception | MessageBox with error message; application continues running | `Log.Error()`, `e.Handled = true` |
| Unhandled domain exception | No UI; logged as fatal | `Log.Fatal()` |
| Unobserved task exception | No UI; logged and observed | `Log.Error()`, `e.SetObserved()` |
| Navigation page type not found | No visible error; page title remains unchanged or shows fallback | Log warning, skip navigation |
| Frame is null during NavigateTo | Navigation silently no-ops | `if (Frame == null) return;` guard in NavigationService |
| MCP server fails to start on startup (`--enable-mcp`) | StatusMessage shows error; application continues without MCP | `Log.Error()`, catch in `StartMCPServerAsync` |
| MCP server fails to stop on shutdown | Logged; shutdown continues | `catch` in `OnExit`, server abandoned |
| GUI proxy timer tick throws | Logged; timer continues | Exception caught within tick handler |
| Window icon file missing | No icon displayed; silently ignored | Try-catch around `iconbitmap` (Python pattern maps to try-catch in C# icon loading) |
| Serilog configuration fails | Console fallback; file logs unavailable | Serilog bootstrapping with minimal config |

## 10. Implementation Notes

### WPF Frame Navigation vs. Python Dynamic Content Replacement

Python replaces the content area by destroying all child widgets and recreating them:
```python
def clear_main_content(self):
    for widget in self.main_content.winfo_children():
        widget.destroy()
    self.create_welcome_screen()
```

The C# port uses WPF `Frame` navigation, which is a proper page-based navigation model:
```csharp
public void NavigateTo(string pageName)
{
    var pageType = Type.GetType($"OdooAutoCAD.App.Views.Pages.{pageName}Page");
    if (pageType != null)
    {
        Frame.Navigate(Activator.CreateInstance(pageType));
    }
}
```

Key differences:
- WPF Frame maintains a navigation journal (`CanGoBack`, `GoBack()`) whereas Python has no back navigation.
- WPF Pages are instantiated via reflection; each page is a full `Page` class with its own XAML and ViewModel.
- `NavigationUIVisibility="Hidden"` suppresses the default WPF navigation chrome (back/forward buttons).
- Python sidebar buttons directly invoke business logic; C# sidebar buttons trigger page navigation, and business logic lives in page ViewModels.

### MVVM Pattern with CommunityToolkit.Mvvm

The shell uses `CommunityToolkit.Mvvm` source generators for clean MVVM:
- `[ObservableProperty]` generates `INotifyPropertyChanged` backing fields for `CurrentPageTitle`, `StatusMessage`, connection status properties.
- `[RelayCommand]` generates `ICommand` implementations for `RefreshAsync`, `StartMCPServerAsync`, `StopMCPServerAsync`.
- `MainViewModel` is resolved from DI in `MainWindow` constructor and set as `DataContext`.
- All UI updates flow through property bindings; no direct UI manipulation in ViewModel code.

### Resource Dictionaries for Theming (App.xaml)

The Python `UITheme` class with dictionaries (`COLORS`, `FONTS`, `SIZES`) maps to WPF `ResourceDictionary` entries in `App.xaml`:

| Python UITheme | WPF App.xaml Resource | Value |
|---|---|---|
| `primary` (#2B579A) | `PrimaryColor` / `PrimaryBrush` | #1A73E8 (adjusted for WPF) |
| `primary_dark` (#1e3a6f) | `PrimaryDarkColor` / `PrimaryDarkBrush` | #1557B0 |
| `success` (#4CAF50) | `SuccessColor` / `SuccessBrush` | #388E3C |
| `warning` (#FF9800) | `WarningColor` / `WarningBrush` | #FFA000 |
| `error` (#F44336) | `ErrorColor` / `ErrorBrush` | #D32F2F |
| `background` (#FFFFFF) | `BackgroundColor` / `BackgroundBrush` | #F5F5F5 |
| `surface` (#F8F9FA) | `SurfaceColor` / `SurfaceBrush` | #FFFFFF |
| `border` (#E0E0E0) | `BorderColor` / `BorderBrush` | #E0E0E0 |
| `text_primary` (#212121) | `TextPrimaryColor` / `TextPrimaryBrush` | #212121 |
| `text_secondary` (#757575) | `TextSecondaryColor` / `TextSecondaryBrush` | #757575 |

Additionally, Python button-specific colors from `form_main_modern.py` map to additional brush resources or inline styles:

| Python Button Context | Hex Value | C# Approach |
|---|---|---|
| Odoo button (disconnected) | #1976D2 | Use `PrimaryBrush` or define `OdooBrush` |
| AutoCAD button (disconnected) | #7B1FA2 | Define `AutoCADPurpleBrush` |
| Connected state | #2E7D32 | Use `SuccessBrush` or define `ConnectedBrush` |
| BOQ orange | #F57C00 | Use `WarningBrush` |
| Danger red | #D32F2F | Use `ErrorBrush` |
| SSE orange | #FF9800 | Define `SSEAccentBrush` |

Consider splitting resources into multiple `ResourceDictionary` files if the number of resources grows:
- `Themes/Colors.xaml` - Color and brush definitions
- `Themes/Styles.xaml` - Control styles (PrimaryButton, NavButton, CardPanel, ModernTextBox)
- `Themes/Fonts.xaml` - FontFamily definitions

### Font System (Microsoft JhengHei UI for CJK)

Python uses `FontManager` with a cross-platform fallback chain. On Windows, the priority is:
1. Microsoft JhengHei UI
2. Microsoft JhengHei
3. SimHei
4. Arial Unicode MS

Font sizes in Python: title=16pt, heading=14pt, body=11pt, small=9pt, button=10pt.

For C#/WPF, set the default font at the Window or Application level:
```xml
<Window FontFamily="Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS"
        FontSize="12">
```

Or define a `FontFamily` resource:
```xml
<FontFamily x:Key="AppFontFamily">Microsoft JhengHei UI, Microsoft JhengHei, SimHei, Arial Unicode MS</FontFamily>
```

WPF handles CJK fallback natively through composite font families. Ensure the font resource is referenced throughout all styles and templates.

### DispatcherTimer for GUI Proxy Polling (100ms)

The Python pattern:
```python
def process_gui_proxy_requests(self):
    processed = self.gui_proxy.process_requests()
    self.after(100, self.process_gui_proxy_requests)
```

Maps to C# in `App.xaml.cs`:
```csharp
_guiProxyTimer = new DispatcherTimer
{
    Interval = TimeSpan.FromMilliseconds(100)
};
_guiProxyTimer.Tick += (s, e) => { guiProxy.ProcessRequests(); };
_guiProxyTimer.Start();
```

Critical implementation notes:
- The timer runs on the WPF dispatcher thread (UI thread), ensuring all COM operations executed through the proxy happen on the STA thread.
- The 100ms interval matches the Python implementation and provides responsive execution without excessive CPU usage.
- The timer must be stopped in `OnExit` before disposing the host.
- `IGUIProxy.Start()` must be called before the timer begins ticking.
- The `GUIProxy` uses `ConcurrentQueue<GUIProxyRequest>` for thread-safe request queuing from MCP server threads.

### DI Container Registration Pattern

`App.xaml.cs` uses `Microsoft.Extensions.Hosting` with `IHost` for DI:

```csharp
_host = Host.CreateDefaultBuilder()
    .UseSerilog()
    .ConfigureServices((context, services) =>
    {
        // 1. Core threading
        services.AddSingleton<IGUIProxy, GUIProxy>();

        // 2. MCP services (factory pattern for controlled construction)
        services.AddSingleton<MCPToolRegistry>(sp => new MCPToolRegistry(...));
        services.AddSingleton<MCPSSEServer>(sp => new MCPSSEServer(...));

        // 3. ViewModels
        services.AddTransient<MainViewModel>();

        // 4. Application services
        services.AddSingleton<INavigationService, NavigationService>();
    })
    .Build();
```

Key patterns:
- Singleton for stateful services (GUIProxy, MCP, Navigation).
- Transient for ViewModels (fresh instance per resolution, though MainViewModel could be singleton for shell state persistence).
- Factory registration for `MCPToolRegistry` and `MCPSSEServer` to control constructor parameters.
- `App.Services` static property for service location in code-behind (used by `MainWindow` constructor).
- Optional service resolution (`sp.GetService<T>()`) for services not yet implemented (IAutoCADService, IOdooService).

### Startup Sequence Mapping: Python to C#

| Step | Python (`ModernFormMain.__init__`) | C# (`App.OnStartup` + `MainWindow()`) |
|------|------|------|
| 1 | `UITheme.setup_theme("system", "blue")` | App.xaml ResourceDictionary loaded automatically |
| 2 | `self.title(...)`, `self.iconbitmap(...)` | `Title` and `Icon` properties in MainWindow.xaml |
| 3 | `setup_window_geometry()` | `Height`, `Width`, `MinHeight`, `MinWidth`, `WindowStartupLocation` in XAML |
| 4 | `setup_window_style()` (WM_DELETE_WINDOW) | `Window.Closing` event or `App.OnExit` |
| 5 | `create_ui()` (top bar, sidebar, content, bottom) | MainWindow.xaml declarative XAML layout |
| 6 | `setup_log_util()` | Serilog configured in `OnStartup` |
| 7 | `init_utilities()` (Odoo, AutoCAD, BOQ, PR, MCP, GUI proxy) | DI `ConfigureServices()` registers all services |
| 8 | `start_gui_proxy_processing()` (100ms timer) | `InitializeGUIProxyTimer()` in `OnStartup` |
| 9 | `auto_start_sse_server()` | `--enable-mcp` flag check in `OnStartup` |
| 10 | `update_connection_status()` | `MainViewModel.UpdateAllStatus()` in constructor |

### Window Chrome Customization

The current C# implementation uses default Windows chrome. To match the Python top bar with colored background:
- Consider using `WindowChrome` to extend client area into the title bar.
- Alternatively, keep standard chrome and use the header bar panel within the content area (current approach).
- If custom chrome is desired, use `WindowStyle="None"` with a custom title bar that includes minimize/maximize/close buttons and drag behavior.
- For v6.0, the recommended approach is to keep standard Windows chrome for native look and feel, and use the header bar for page-contextual information.

### Color Scheme Mapping from Python to WPF Resources

Python `form_main_modern.py` uses inline hex colors for specific buttons. These should be consolidated into named WPF resources for maintainability:

```xml
<!-- Application-specific accent colors (supplement to core palette) -->
<Color x:Key="OdooBlueColor">#1976D2</Color>
<Color x:Key="AutoCADPurpleColor">#7B1FA2</Color>
<Color x:Key="ConnectedGreenColor">#2E7D32</Color>
<Color x:Key="BOQOrangeColor">#F57C00</Color>
<Color x:Key="DangerRedColor">#D32F2F</Color>
<Color x:Key="SSEOrangeColor">#FF9800</Color>
<Color x:Key="ToolsBrownColor">#5D4037</Color>

<SolidColorBrush x:Key="OdooBlueBrush" Color="{StaticResource OdooBlueColor}"/>
<SolidColorBrush x:Key="AutoCADPurpleBrush" Color="{StaticResource AutoCADPurpleColor}"/>
<SolidColorBrush x:Key="ConnectedGreenBrush" Color="{StaticResource ConnectedGreenColor}"/>
<!-- ... etc ... -->
```

### Navigation Button Highlighting for Active Page

Python does not highlight the active sidebar button (it uses direct action invocation). The C# implementation must add active state tracking:

1. **ViewModel approach**: `MainViewModel.ActiveNavButton` property holds the current button name (e.g., "BtnDashboard"). Each button binds its background to a converter that checks if its name matches `ActiveNavButton`.

2. **Style trigger approach**: Define a `Tag` or attached property on each button. Use a `DataTrigger` in the `NavButton` style:
```xml
<DataTrigger Binding="{Binding ActiveNavButton}" Value="BtnDashboard">
    <Setter Property="Background" Value="#E3F2FD"/>
    <Setter TargetName="border" Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
</DataTrigger>
```

3. **Code-behind approach** (current implementation): In `NavButton_Click`, iterate all nav buttons to reset styles, then highlight the clicked button. This is simpler but less MVVM-pure.

The recommended approach is (1) or (2) for MVVM compliance. The `NavigateTo` method in MainWindow code-behind should set `_viewModel.ActiveNavButton` alongside `_viewModel.CurrentPageTitle`.
