# Master Task Sequence — Odoo-AutoCAD C# WPF

> **Total Tasks**: 500 | **Phases**: 4 | **Sprints**: 11
> **Estimates**: 313S + 160M + 27L
> **User Stories**: 78 | **Generated**: 2026-02-06

## Overview

This document sequences all implementation tasks for the Odoo-AutoCAD C# WPF desktop application
across 4 phases and 11 sprints. Tasks are ordered by dependency within each sprint, with
foundational infrastructure first and feature-specific work building on top. Each task references
its parent User Story for full context and acceptance criteria.

## Progress Summary

| Phase | Sprints | User Stories | Tasks | S | M | L | Status |
|-------|---------|-------------|-------|---|---|---|--------|
| Phase 1: Core Infrastructure | 1–3 | 17 | 103 | 57 | 39 | 7 | Not Started |
| Phase 2: Business Logic | 4–6 | 17 | 121 | 71 | 38 | 12 | Not Started |
| Phase 3: Feature Completion | 7–9 | 30 | 197 | 138 | 55 | 4 | Not Started |
| Phase 4: AI Integration | 10–11 | 14 | 79 | 47 | 28 | 4 | Not Started |

---

## Phase 1: Core Infrastructure

### Sprint 1: App Shell & Navigation
> **Focus**: Application skeleton, sidebar navigation, window management, fonts, clean shutdown
> **User Stories**: US-008-01, US-008-02, US-008-05, US-008-08, US-008-09, US-008-10
> **Tasks**: 32 (22S + 9M + 1L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 1 | TASK-008-01-01 | Define sidebar XAML layout with title area, navigation buttons, and... | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | L | — | [ ] |
| 2 | TASK-008-01-02 | Implement NavButton style with transparent background, left-align, ... | `...ainWindow.xaml` (Window.Resources) or `App.xaml` | M | TASK-008-01-01 | [ ] |
| 3 | TASK-008-01-03 | Create INavigationService interface with NavigateTo, GoBack, CanGoB... | `src/OdooAutoCAD.App/Services/NavigationService.cs` | M | — | [ ] |
| 4 | TASK-008-01-04 | Implement NavigationService with reflection-based page resolution | `src/OdooAutoCAD.App/Services/NavigationService.cs` | M | TASK-008-01-03 | [ ] |
| 5 | TASK-008-01-05 | Add NavigateCommand and page title mapping to MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | M | TASK-008-01-03 | [ ] |
| 6 | TASK-008-01-06 | Wire MainWindow constructor to set Frame on NavigationService and n... | `src/OdooAutoCAD.App/Views/MainWindow.xaml.cs` | S | TASK-008-01-04, TASK-008-01-05 | [ ] |
| 7 | TASK-008-01-07 | Register INavigationService as singleton in DI container | `src/OdooAutoCAD.App/App.xaml.cs` | S | TASK-008-01-03 | [ ] |
| 8 | TASK-008-02-01 | Add ActiveNavButton observable property to MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | — | [ ] |
| 9 | TASK-008-02-02 | Implement NavButton style DataTriggers for active state highlighting | `` | M | TASK-008-02-01 | [ ] |
| 10 | TASK-008-02-03 | Update NavigateCommand handler to set ActiveNavButton alongside Cur... | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | TASK-008-02-01 | [ ] |
| 11 | TASK-008-02-04 | Assign Tag property to each sidebar button matching ActiveNavButton... | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | TASK-008-02-02 | [ ] |
| 12 | TASK-008-05-01 | Define MainWindow XAML root element with size, position, and backgr... | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | — | [ ] |
| 13 | TASK-008-05-02 | Define two-column Grid with fixed sidebar and star-sized content co... | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | TASK-008-05-01 | [ ] |
| 14 | TASK-008-05-03 | Define three-row content area Grid | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | TASK-008-05-01 | [ ] |
| 15 | TASK-008-05-04 | Define BackgroundBrush resource in application ResourceDictionary | `src/OdooAutoCAD.App/App.xaml` | S | — | [ ] |
| 16 | TASK-008-05-05 | Set window Title and Icon | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | — | [ ] |
| 17 | TASK-008-08-01 | Define header bar XAML with page title TextBlock and action buttons | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | M | — | [ ] |
| 18 | TASK-008-08-02 | Add CurrentPageTitle observable property to MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | — | [ ] |
| 19 | TASK-008-08-03 | Add RefreshCommand to MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | — | [ ] |
| 20 | TASK-008-08-04 | Define page title display name mapping | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | TASK-008-08-02 | [ ] |
| 21 | TASK-008-08-05 | Apply SurfaceBrush background and bottom border styling to header bar | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | TASK-008-08-01 | [ ] |
| 22 | TASK-008-09-01 | Implement App.OnExit with ordered cleanup sequence | `src/OdooAutoCAD.App/App.xaml.cs` | M | — | [ ] |
| 23 | TASK-008-09-02 | Register DispatcherUnhandledException handler | `src/OdooAutoCAD.App/App.xaml.cs` | S | — | [ ] |
| 24 | TASK-008-09-03 | Register AppDomain.UnhandledException handler | `src/OdooAutoCAD.App/App.xaml.cs` | S | — | [ ] |
| 25 | TASK-008-09-04 | Register TaskScheduler.UnobservedTaskException handler | `src/OdooAutoCAD.App/App.xaml.cs` | S | — | [ ] |
| 26 | TASK-008-09-05 | Add try-catch around MCP server stop to prevent shutdown blocking | `src/OdooAutoCAD.App/App.xaml.cs` | S | TASK-008-09-01 | [ ] |
| 27 | TASK-008-09-06 | Verify no port locks remain after shutdown via integration test | `` | M | TASK-008-09-01, TASK-008-09-05 | [ ] |
| 28 | TASK-008-10-01 | Define AppFontFamily resource with fallback chain in ResourceDictio... | `src/OdooAutoCAD.App/App.xaml` | S | — | [ ] |
| 29 | TASK-008-10-02 | Set FontFamily at the Window level referencing AppFontFamily | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | TASK-008-10-01 | [ ] |
| 30 | TASK-008-10-03 | Merge font resources into App.xaml ResourceDictionary | `src/OdooAutoCAD.App/App.xaml` | S | TASK-008-10-01 | [ ] |
| 31 | TASK-008-10-04 | Verify all named styles reference AppFontFamily | `src/OdooAutoCAD.App/App.xaml` | S | TASK-008-10-01 | [ ] |
| 32 | TASK-008-10-05 | Test CJK rendering with sample Chinese strings across all UI areas | `` | M | TASK-008-10-02, TASK-008-10-04 | [ ] |

### Sprint 2: Status Bar, Log Panel & Dashboard
> **Focus**: Connection status bar, system log panel, auto-start services, dashboard views
> **User Stories**: US-008-03, US-008-04, US-008-06, US-001-01, US-001-02
> **Tasks**: 31 (15S + 15M + 1L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 33 | TASK-008-03-01 | Add connection status observable properties to MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | M | — | [ ] |
| 34 | TASK-008-03-02 | Define connection status XAML section in sidebar with Ellipse indic... | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | M | TASK-008-03-01 | [ ] |
| 35 | TASK-008-03-03 | Implement status update methods in MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | M | TASK-008-03-01 | [ ] |
| 36 | TASK-008-03-04 | Define connected/disconnected/error color resources | `src/OdooAutoCAD.App/App.xaml` | S | — | [ ] |
| 37 | TASK-008-03-05 | Wire MCP server status to display port number when running | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | TASK-008-03-01, TASK-008-03-03 | [ ] |
| 38 | TASK-008-04-01 | Define status bar XAML with three-column layout | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | — | [ ] |
| 39 | TASK-008-04-02 | Add StatusMessage, CurrentDateTime, and version observable properti... | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | — | [ ] |
| 40 | TASK-008-04-03 | Implement collapsible log panel XAML with ScrollViewer | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | M | TASK-008-04-01 | [ ] |
| 41 | TASK-008-04-04 | Create Serilog sink or log collection for UI panel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | M | TASK-008-04-03 | [ ] |
| 42 | TASK-008-04-05 | Implement auto-scroll behavior on log panel ScrollViewer | `src/OdooAutoCAD.App/Views/MainWindow.xaml.cs` | S | TASK-008-04-03, TASK-008-04-04 | [ ] |
| 43 | TASK-008-04-06 | Add DispatcherTimer to update CurrentDateTime every second | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | TASK-008-04-02 | [ ] |
| 44 | TASK-008-06-01 | Implement App.OnStartup with Serilog configuration, DI host build, ... | `src/OdooAutoCAD.App/App.xaml.cs` | L | — | [ ] |
| 45 | TASK-008-06-02 | Configure DI services registration in correct order | `src/OdooAutoCAD.App/App.xaml.cs` | M | TASK-008-06-01 | [ ] |
| 46 | TASK-008-06-03 | Implement --enable-mcp flag parsing with case-insensitive check | `src/OdooAutoCAD.App/App.xaml.cs` | S | TASK-008-06-01 | [ ] |
| 47 | TASK-008-06-04 | Initialize DispatcherTimer for GUI proxy polling after DI host start | `src/OdooAutoCAD.App/App.xaml.cs` | M | TASK-008-06-01 | [ ] |
| 48 | TASK-008-06-05 | Implement error handling for MCP server startup failure | `src/OdooAutoCAD.App/App.xaml.cs` | S | TASK-008-06-03 | [ ] |
| 49 | TASK-008-06-06 | Implement IGUIProxy interface with Start, Stop, ProcessRequests met... | `src/Threading/IGUIProxy.cs` | M | — | [ ] |
| 50 | TASK-001-01-01 | Create DashboardPage XAML layout with connection status cards (Auto... | `Views/Pages/DashboardPage.xaml` | M | — | [ ] |
| 51 | TASK-001-01-02 | Add DashboardViewModel properties for connection status | `ViewModels/DashboardViewModel.cs` | M | — | [ ] |
| 52 | TASK-001-01-03 | Implement BoolToColorConverter for green/red-gray status indicator ... | `Views/Pages/DashboardPage.xaml` | S | TASK-001-01-01 | [ ] |
| 53 | TASK-001-01-04 | Subscribe DashboardViewModel to IAutoCADService and IOdooService st... | `ViewModels/DashboardViewModel.cs` | M | TASK-001-01-02 | [ ] |
| 54 | TASK-001-01-05 | Register DashboardViewModel as singleton in DI container to maintai... | `ViewModels/DashboardViewModel.cs` | S | TASK-001-01-02 | [ ] |
| 55 | TASK-001-01-06 | Implement error message display in status cards when connection att... | `ViewModels/DashboardViewModel.cs` | S | TASK-001-01-02 | [ ] |
| 56 | TASK-001-02-01 | Add connect buttons inside status cards in XAML | `Views/Pages/DashboardPage.xaml` | S | — | [ ] |
| 57 | TASK-001-02-02 | Implement ConnectAutoCADCommand with async execution | `ViewModels/DashboardViewModel.cs` | M | — | [ ] |
| 58 | TASK-001-02-03 | Implement ConnectOdooCommand with async execution | `ViewModels/DashboardViewModel.cs` | M | — | [ ] |
| 59 | TASK-001-02-04 | Add CanExecute logic to connect commands | `ViewModels/DashboardViewModel.cs` | S | TASK-001-02-02, TASK-001-02-03 | [ ] |
| 60 | TASK-001-02-05 | Create Quick Actions section in XAML with workflow shortcut buttons | `Views/Pages/DashboardPage.xaml` | M | TASK-001-02-01 | [ ] |
| 61 | TASK-001-02-06 | Implement navigation commands using INavigationService | `ViewModels/DashboardViewModel.cs` | M | — | [ ] |
| 62 | TASK-001-02-07 | Add CanExecute validation for workflow shortcuts | `ViewModels/DashboardViewModel.cs` | S | TASK-001-02-06 | [ ] |
| 63 | TASK-001-02-08 | Re-evaluate command CanExecute states when connection status changes | `ViewModels/DashboardViewModel.cs` | S | TASK-001-02-04, TASK-001-02-07 | [ ] |

### Sprint 3: AutoCAD & Odoo Core Connections
> **Focus**: AutoCAD connect/layouts/parameter extraction, Odoo credentials/test/appsettings config
> **User Stories**: US-002-01, US-002-02, US-002-03, US-003-01, US-003-02, US-003-10
> **Tasks**: 40 (20S + 15M + 5L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 64 | TASK-002-01-01 | Create AutoCAD page XAML with connection panel | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | M | — | [ ] |
| 65 | TASK-002-01-02 | Implement ConnectCommand and DisconnectCommand in ViewModel | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | TASK-002-01-04, TASK-002-01-05 | [ ] |
| 66 | TASK-002-01-03 | Implement ConnectAsync in AutoCADService with retry logic | `src/AutoCAD/AutoCADService.cs` | L | TASK-002-01-04 | [ ] |
| 67 | TASK-002-01-04 | Define ConnectAsync and DisconnectAsync in IAutoCADService interface | `src/AutoCAD/IAutoCADService.cs` | S | — | [ ] |
| 68 | TASK-002-01-05 | Wire COM operations through IGUIProxy for STA thread safety | `src/Threading/IGUIProxy.cs` | M | TASK-002-01-03 | [ ] |
| 69 | TASK-002-01-06 | Implement connection status properties in ViewModel | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | TASK-002-01-01 | [ ] |
| 70 | TASK-002-02-01 | Add Layouts ListView panel to the AutoCAD page XAML | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | M | — | [ ] |
| 71 | TASK-002-02-02 | Implement Layouts ObservableCollection and SelectedLayout property ... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | — | [ ] |
| 72 | TASK-002-02-03 | Implement GetLayouts() in AutoCADService excluding "Model" | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 73 | TASK-002-02-04 | Implement GetActiveLayout() and SetActiveLayout(name) in AutoCADSer... | `src/AutoCAD/AutoCADService.cs` | M | TASK-002-02-03 | [ ] |
| 74 | TASK-002-02-05 | Define GetLayouts(), GetActiveLayout() in IAutoCADService interface | `src/AutoCAD/IAutoCADService.cs` | S | — | [ ] |
| 75 | TASK-002-02-06 | Implement SelectLayoutCommand to load layout details on selection c... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | TASK-002-02-01, TASK-002-02-02 | [ ] |
| 76 | TASK-002-02-07 | Add visual highlighting for the active layout in the ListView | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | S | TASK-002-02-01 | [ ] |
| 77 | TASK-002-03-01 | Add Layout Details panel and Table Data DataGrid to XAML | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | M | — | [ ] |
| 78 | TASK-002-03-02 | Implement ExtractParametersCommand in ViewModel | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | TASK-002-03-08 | [ ] |
| 79 | TASK-002-03-03 | Implement GetLayoutsValues() and GetLayoutValues(name) in AutoCADSe... | `src/AutoCAD/AutoCADService.cs` | L | TASK-002-03-04, TASK-002-03-05 | [ ] |
| 80 | TASK-002-03-04 | Implement get_attribute_values equivalent for block attribute extra... | `src/AutoCAD/AutoCADService.cs` | L | — | [ ] |
| 81 | TASK-002-03-05 | Implement get_table_data equivalent with validation | `src/AutoCAD/AutoCADService.cs` | L | TASK-002-03-06, TASK-002-03-07 | [ ] |
| 82 | TASK-002-03-06 | Implement LM_UnFormat equivalent in C# | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 83 | TASK-002-03-07 | Implement empty row filtering logic | `src/AutoCAD/AutoCADService.cs` | S | — | [ ] |
| 84 | TASK-002-03-08 | Define TableRowData model class and LayoutAttributes in ViewModel | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | — | [ ] |
| 85 | TASK-002-03-09 | Bind DataGrid columns to TableRowData properties | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | M | TASK-002-03-01 | [ ] |
| 86 | TASK-003-01-01 | Create Connection Settings GroupBox with credential input fields | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | L | — | [ ] |
| 87 | TASK-003-01-02 | Add ViewModel properties for credential fields with change notifica... | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 88 | TASK-003-01-03 | Implement INotifyDataErrorInfo validation for credential fields | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-01-02 | [ ] |
| 89 | TASK-003-01-04 | Load default values from IConfiguration on ViewModel initialization | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-01-02 | [ ] |
| 90 | TASK-003-01-05 | Create PasswordBox attached behavior for MVVM binding | `src/OdooAutoCAD.App/Helpers/PasswordBoxHelper.cs` | M | TASK-003-01-01 | [ ] |
| 91 | TASK-003-01-06 | Bind Connect/Test Connection button IsEnabled to validation state | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-01-01, TASK-003-01-03 | [ ] |
| 92 | TASK-003-02-01 | Add "Test Connection" button to XAML button bar | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 93 | TASK-003-02-02 | Implement TestConnectionCommand in ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-02-03 | [ ] |
| 94 | TASK-003-02-03 | Implement TestConnectionAsync in OdooService (non-persistent) | `src/Odoo/OdooService.cs` | M | TASK-003-02-04 | [ ] |
| 95 | TASK-003-02-04 | Define TestConnectionAsync overload in IOdooService interface | `src/Odoo/IOdooService.cs` | S | — | [ ] |
| 96 | TASK-003-02-05 | Add error handling with user-friendly messages for each failure sce... | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-02-02 | [ ] |
| 97 | TASK-003-02-06 | Bind button IsEnabled to validation state and IsRunning | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-02-01 | [ ] |
| 98 | TASK-003-10-01 | Add Odoo section to appsettings.json | `src/OdooAutoCAD.App/appsettings.json` | S | — | [ ] |
| 99 | TASK-003-10-02 | Verify OdooSettings POCO class for strongly-typed binding | `...dooAutoCAD.Configuration/ConfigurationLoader.cs` | S | — | [ ] |
| 100 | TASK-003-10-03 | Inject IConfiguration and bind Odoo section in ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-10-01, TASK-003-10-02 | [ ] |
| 101 | TASK-003-10-04 | Pre-populate credential fields from configuration | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-10-03 | [ ] |
| 102 | TASK-003-10-05 | Configure HttpClient.Timeout from TimeoutSeconds | `src/Odoo/OdooService.cs` | S | — | [ ] |
| 103 | TASK-003-10-06 | Handle missing Odoo section gracefully | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |

---

## Phase 2: Business Logic

### Sprint 4: Odoo Advanced & BOQ Foundation
> **Focus**: Odoo connection status/disconnect/sync product catalog/persist credentials, BOQ extract/review
> **User Stories**: US-003-03, US-003-04, US-003-05, US-003-12, US-004-01, US-004-02
> **Tasks**: 42 (26S + 10M + 6L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 104 | TASK-003-03-01 | Create Connection Status panel in XAML with colored indicators | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | M | — | [ ] |
| 105 | TASK-003-03-02 | Implement IsConnected property with change notification | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 106 | TASK-003-03-03 | Implement ConnectionStatusText and display properties | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-03-02 | [ ] |
| 107 | TASK-003-03-04 | Create BoolToColorConverter for status indicator binding | `...oAutoCAD.App/Converters/BoolToColorConverter.cs` | S | TASK-003-03-01 | [ ] |
| 108 | TASK-003-03-05 | Implement IsConnected property check in OdooService | `src/Odoo/OdooService.cs` | S | — | [ ] |
| 109 | TASK-003-03-06 | Add connection state change event for cross-page reactivity | `src/Odoo/IOdooService.cs` | M | — | [ ] |
| 110 | TASK-003-04-01 | Add "Disconnect" button to XAML button bar | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 111 | TASK-003-04-02 | Implement DisconnectCommand in ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-04-03 | [ ] |
| 112 | TASK-003-04-03 | Implement DisconnectAsync in OdooService with full state cleanup | `src/Odoo/OdooService.cs` | L | — | [ ] |
| 113 | TASK-003-04-04 | Reset ViewModel status properties on disconnect | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-04-02 | [ ] |
| 114 | TASK-003-04-05 | Bind Disconnect button IsEnabled to IsConnected | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-04-01 | [ ] |
| 115 | TASK-003-04-06 | Fire connection state change event on disconnect | `src/Odoo/OdooService.cs` | S | — | [ ] |
| 116 | TASK-003-05-01 | Create Product Catalog section in XAML with DataGrid and controls | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | L | — | [ ] |
| 117 | TASK-003-05-02 | Implement SyncProductsCommand in ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | — | [ ] |
| 118 | TASK-003-05-03 | Add Products collection and sync status properties to ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 119 | TASK-003-05-04 | Implement local product caching | `src/OdooAutoCAD.Data/Context/AppDbContext.cs` | M | — | [ ] |
| 120 | TASK-003-05-05 | Configure DataGrid with VirtualizingStackPanel | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-05-01 | [ ] |
| 121 | TASK-003-05-06 | Implement sync result summary display | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-05-02 | [ ] |
| 122 | TASK-003-05-07 | Add SyncHistory entry after product sync | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-05-06 | [ ] |
| 123 | TASK-003-12-01 | Add "Remember Me" checkbox to XAML | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 124 | TASK-003-12-02 | Create ICredentialService interface | `src/Services/ICredentialService.cs` | S | — | [ ] |
| 125 | TASK-003-12-03 | Implement CredentialService using Windows DPAPI | `src/Services/DpapiCredentialService.cs` | L | TASK-003-12-02 | [ ] |
| 126 | TASK-003-12-04 | Alternative: Implement CredentialService using Windows Credential M... | `src/Services/WindowsCredentialService.cs` | L | TASK-003-12-02 | [ ] |
| 127 | TASK-003-12-05 | Load persisted credentials on ViewModel initialization | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-12-03 or TASK-003-12-04 | [ ] |
| 128 | TASK-003-12-06 | Save credentials on successful connection | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-12-05 | [ ] |
| 129 | TASK-003-12-07 | Clear persisted credentials when "Remember Me" is unchecked | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-12-06 | [ ] |
| 130 | TASK-003-12-08 | Add error handling for credential retrieval failures | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 131 | TASK-004-01-01 | Add "Extract from AutoCAD" button to BOQ Manager page with binding ... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 132 | TASK-004-01-02 | Implement ExtractCommand in ViewModel that calls IGUIProxy for COM ... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | L | TASK-004-01-01 | [ ] |
| 133 | TASK-004-01-03 | Implement GetLayoutsValues() in IAutoCADService to iterate layouts,... | `src/AutoCAD/AutoCADService.cs` | L | TASK-004-01-07 | [ ] |
| 134 | TASK-004-01-04 | Implement MTextFormatter.UnFormat() static method for stripping MTe... | `src/Utilities/MTextFormatter.cs` | M | — | [ ] |
| 135 | TASK-004-01-05 | Register extraction handler in IGUIProxy for STA thread execution | `src/OdooAutoCAD.App/App.xaml.cs` | M | TASK-004-01-03 | [ ] |
| 136 | TASK-004-01-06 | Add progress reporting with IProgress<T> pattern during layout iter... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-01-02 | [ ] |
| 137 | TASK-004-01-07 | Implement chk_legal_table equivalent: validate 9 columns and HEADER... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 138 | TASK-004-01-08 | Implement row skip logic for empty qty + product_no rows | `src/BOQ/BOQProcessor.cs` | S | TASK-004-01-04 | [ ] |
| 139 | TASK-004-02-01 | Define DataGrid XAML with all columns and CollectionViewSource grou... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | M | — | [ ] |
| 140 | TASK-004-02-02 | Implement GroupStyle with collapsible Expander showing layout metad... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | M | TASK-004-02-01 | [ ] |
| 141 | TASK-004-02-03 | Implement summary panel using UniformGrid with 5 cells and colored ... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 142 | TASK-004-02-04 | Create ValidationStatusToIconConverter for Status column DataTemplate | `...p/Converters/ValidationStatusToIconConverter.cs` | S | — | [ ] |
| 143 | TASK-004-02-05 | Bind AllEntries ObservableCollection to DataGrid ItemsSource with g... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-02-01 | [ ] |
| 144 | TASK-004-02-06 | Implement RowValidationStatusToBackgroundConverter for row-level co... | `...ers/RowValidationStatusToBackgroundConverter.cs` | S | — | [ ] |
| 145 | TASK-004-02-07 | Wire summary count properties with OnPropertyChanged notifications | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |

### Sprint 5: BOQ Processing Pipeline
> **Focus**: BOQ validate against products, push to Odoo, ID writeback, skipped rows, summary
> **User Stories**: US-004-03, US-004-04, US-004-05, US-004-06, US-004-07
> **Tasks**: 33 (20S + 9M + 4L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 146 | TASK-004-03-01 | Add "Validate All" button to BOQ Manager page bound to ValidateCommand | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 147 | TASK-004-03-02 | Implement ValidateCommand that calls IBOQProcessor.ValidateBOQAsync... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-03-01 | [ ] |
| 148 | TASK-004-03-03 | Implement ValidateBOQAsync with product_no resolution (local cache ... | `src/BOQ/BOQProcessor.cs` | L | — | [ ] |
| 149 | TASK-004-03-04 | Implement qty validation rules: positive numeric required, zero=War... | `src/BOQ/BOQProcessor.cs` | M | — | [ ] |
| 150 | TASK-004-03-05 | Implement project context validation via IOdooService.GetProjectAsy... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 151 | TASK-004-03-06 | Build BOQValidationResult/BOQValidationError model population with ... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 152 | TASK-004-03-07 | Update BOQEntryRow.ValidationStatus and ValidationMessage after val... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-03-02 | [ ] |
| 153 | TASK-004-03-08 | Recalculate summary counts (ValidItems, WarningItems, ErrorItems) p... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-03-02 | [ ] |
| 154 | TASK-004-04-01 | Add "Push to Odoo" button with CanExecute binding to validation sta... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 155 | TASK-004-04-02 | Implement PushToOdooCommand with pre-validation check and IBOQProce... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | L | TASK-004-04-01 | [ ] |
| 156 | TASK-004-04-03 | Implement ImportToBOQAsync in IOdooService (equivalent to import2bo... | `src/Odoo/OdooService.cs` | L | — | [ ] |
| 157 | TASK-004-04-04 | Parse Odoo response to extract header_id and detail per layout for ... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-04-02 | [ ] |
| 158 | TASK-004-04-05 | Add push progress reporting with IProgress<T> pattern | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-04-02 | [ ] |
| 159 | TASK-004-04-06 | Add push result summary panel XAML with status text and last push info | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 160 | TASK-004-04-07 | Handle Odoo error_code responses and display user-facing error mess... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-04-02 | [ ] |
| 161 | TASK-004-04-08 | Handle partial failure and network timeout scenarios with retry cap... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-04-02 | [ ] |
| 162 | TASK-004-05-01 | Implement set_layouts_tables_id equivalent via IGUIProxy for writin... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | L | TASK-004-05-02, TASK-004-05-03 | [ ] |
| 163 | TASK-004-05-02 | Implement SetTableValue() in IAutoCADService for writing individual... | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 164 | TASK-004-05-03 | Register writeback handler in IGUIProxy for STA thread execution | `src/OdooAutoCAD.App/App.xaml.cs` | M | — | [ ] |
| 165 | TASK-004-05-04 | Implement get_detail_id_index equivalent: match product_no to detai... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 166 | TASK-004-05-05 | Update BOQEntryRow.DetailId in the DataGrid after successful writeback | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-05-01 | [ ] |
| 167 | TASK-004-05-06 | Handle per-layout writeback errors with user-facing messages and re... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-05-01 | [ ] |
| 168 | TASK-004-06-01 | Track skipped row count during extraction (empty qty + product_no r... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 169 | TASK-004-06-02 | Track skipped table count during extraction (illegal tables) | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 170 | TASK-004-06-03 | Display SkippedItems in summary panel with gray background | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 171 | TASK-004-06-04 | Add tooltip on Skipped count showing breakdown (empty rows vs illeg... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | TASK-004-06-03 | [ ] |
| 172 | TASK-004-06-05 | Store per-layout skip information for user reference | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-06-01, TASK-004-06-02 | [ ] |
| 173 | TASK-004-07-01 | Implement summary bar XAML using UniformGrid with 5 cells and color... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 174 | TASK-004-07-02 | Bind summary count properties to UI elements | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | TASK-004-07-01 | [ ] |
| 175 | TASK-004-07-03 | Implement summary calculation method that tallies counts from AllEn... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |
| 176 | TASK-004-07-04 | Trigger summary recalculation after extraction completes | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-07-03 | [ ] |
| 177 | TASK-004-07-05 | Trigger summary recalculation after validation completes | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-07-03 | [ ] |
| 178 | TASK-004-07-06 | Add LayoutsFound count property and bind to summary panel | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |

### Sprint 6: Purchase Requisition & Settings Core
> **Focus**: PR conversion/list/submit, Settings Odoo connection/environments/test connection
> **User Stories**: US-005-01, US-005-02, US-005-04, US-007-01, US-007-02, US-007-03
> **Tasks**: 46 (25S + 19M + 2L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 179 | TASK-005-01-01 | Add "Convert BOQ to PR" button to page layout | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 180 | TASK-005-01-02 | Implement ConvertBOQToPRCommand in ViewModel | `ViewModels/PurchaseRequisitionViewModel.cs` | M | TASK-005-01-01, TASK-005-01-03 | [ ] |
| 181 | TASK-005-01-03 | Implement ConvertBOQToPRAsync in Odoo service | `Odoo/IOdooService.cs`, `Odoo/OdooService.cs` | M | — | [ ] |
| 182 | TASK-005-01-04 | Add optional header_id extraction path via IAutoCADService | `...IAutoCADService.cs`, `AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 183 | TASK-005-01-05 | Wrap AutoCAD COM calls through IGUIProxy for thread safety | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-01-04 | [ ] |
| 184 | TASK-005-01-06 | Add progress indicator and conversion state management | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-01-01, TASK-005-01-02 | [ ] |
| 185 | TASK-005-01-07 | Implement post-conversion PR list refresh and highlighting | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-01-02 | [ ] |
| 186 | TASK-005-01-08 | Implement error handling for conversion failures | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-01-02 | [ ] |
| 187 | TASK-005-01-09 | Add BOQ entry selection UI for partial conversion | `Views/Pages/PurchaseRequisitionPage.xaml` | M | TASK-005-01-01 | [ ] |
| 188 | TASK-005-02-01 | Create PR list ListView layout | `Views/Pages/PurchaseRequisitionPage.xaml` | M | — | [ ] |
| 189 | TASK-005-02-02 | Implement ObservableCollection and data binding | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-02-01 | [ ] |
| 190 | TASK-005-02-03 | Implement GetPurchaseRequisitionsAsync in Odoo service | `Odoo/IOdooService.cs`, `Odoo/OdooService.cs` | M | — | [ ] |
| 191 | TASK-005-02-04 | Implement auto-load on page open | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-02-02, TASK-005-02-03 | [ ] |
| 192 | TASK-005-02-05 | Add Refresh button | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-02-01 | [ ] |
| 193 | TASK-005-02-06 | Implement empty state UI | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-02-01 | [ ] |
| 194 | TASK-005-02-07 | Add IsLoading state with loading indicator | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 195 | TASK-005-02-08 | Implement error handling for PR list fetch failures | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 196 | TASK-005-02-09 | Map PREntry to PRDisplayItem with computed properties | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-02-02 | [ ] |
| 197 | TASK-005-04-01 | Add "Submit for Approval" button to PR detail panel | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 198 | TASK-005-04-02 | Implement SubmitPRCommand with CanExecute logic | `ViewModels/PurchaseRequisitionViewModel.cs` | M | TASK-005-04-01, TASK-005-04-03 | [ ] |
| 199 | TASK-005-04-03 | Implement SubmitPRAsync in Odoo service | `Odoo/IOdooService.cs`, `Odoo/OdooService.cs` | M | — | [ ] |
| 200 | TASK-005-04-04 | Implement confirmation dialog before submission | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-04-02 | [ ] |
| 201 | TASK-005-04-05 | Update PR state in UI after successful submission | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-04-02 | [ ] |
| 202 | TASK-005-04-06 | Implement error handling for submission failures | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-04-02 | [ ] |
| 203 | TASK-005-04-07 | Add loading indicator on Submit button during submission | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-04-01 | [ ] |
| 204 | TASK-007-01-01 | Create SettingsPage XAML with TabControl and Connection tab contain... | `Views/Pages/SettingsPage.xaml` | M | — | [ ] |
| 205 | TASK-007-01-02 | Implement PasswordBox/TextBox toggle for API Token with show/hide b... | `Views/Pages/SettingsPage.xaml` | S | TASK-007-01-01 | [ ] |
| 206 | TASK-007-01-03 | Add SettingsViewModel with Odoo connection properties | `ViewModels/SettingsViewModel.cs` | M | — | [ ] |
| 207 | TASK-007-01-04 | Implement INotifyDataErrorInfo validation for Odoo connection fields | `ViewModels/SettingsViewModel.cs` | M | TASK-007-01-03 | [ ] |
| 208 | TASK-007-01-05 | Implement SaveSettingsCommand persisting to appsettings.json and Se... | `ViewModels/SettingsViewModel.cs` | L | TASK-007-01-03, TASK-007-01-04 | [ ] |
| 209 | TASK-007-01-06 | Extend ConfigurationLoader with SaveToJson method excluding sensiti... | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M | — | [ ] |
| 210 | TASK-007-01-07 | Implement HasUnsavedChanges tracking with original-values snapshot | `ViewModels/SettingsViewModel.cs` | L | TASK-007-01-03, TASK-007-01-01 | [ ] |
| 211 | TASK-007-01-08 | Add SyncLog audit entry on settings save | `ViewModels/SettingsViewModel.cs` | S | TASK-007-01-05 | [ ] |
| 212 | TASK-007-02-01 | Add environment selector dropdown to Connection tab in XAML | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 213 | TASK-007-02-02 | Add SelectedEnvironment, AvailableEnvironments, and ChangeEnvironme... | `ViewModels/SettingsViewModel.cs` | M | — | [ ] |
| 214 | TASK-007-02-03 | Extend ConfigurationLoader.GetAvailableConfigs() to discover appset... | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M | — | [ ] |
| 215 | TASK-007-02-04 | Implement confirmation dialog with change preview and cancel handling | `ViewModels/SettingsViewModel.cs` | M | TASK-007-02-02 | [ ] |
| 216 | TASK-007-02-05 | Load environment-specific values and populate ViewModel fields on c... | `ViewModels/SettingsViewModel.cs` | M | TASK-007-02-03, TASK-007-02-04 | [ ] |
| 217 | TASK-007-02-06 | Persist selected environment to UserPreferences table | `ViewModels/SettingsViewModel.cs` | S | TASK-007-02-02 | [ ] |
| 218 | TASK-007-03-01 | Add "Test Connection" button with inline result display area to Con... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 219 | TASK-007-03-02 | Implement TestConnectionCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M | — | [ ] |
| 220 | TASK-007-03-03 | Add IsTestingConnection, ConnectionTestResult, and ConnectionTestMe... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 221 | TASK-007-03-04 | Implement IOdooService.TestConnectionAsync method | `` | M | — | [ ] |
| 222 | TASK-007-03-05 | Add error classification logic for timeout, auth failure, and netwo... | `ViewModels/SettingsViewModel.cs` | M | TASK-007-03-02 | [ ] |
| 223 | TASK-007-03-06 | Add XAML data triggers for success (green), failure (red), and test... | `Views/Pages/SettingsPage.xaml` | S | TASK-007-03-01, TASK-007-03-03 | [ ] |
| 224 | TASK-007-03-07 | Clear test result when connection fields are modified | `ViewModels/SettingsViewModel.cs` | S | TASK-007-03-02 | [ ] |

---

## Phase 3: Feature Completion

### Sprint 7: AutoCAD P2 & Odoo Search/Filter
> **Focus**: AutoCAD drawing info/PR project/clear IDs/COM monitor, Odoo search/filter/sync time/server info
> **User Stories**: US-002-04, US-002-05, US-002-06, US-002-07, US-003-06, US-003-07, US-003-08, US-003-09, US-003-11
> **Tasks**: 57 (31S + 24M + 2L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 225 | TASK-002-04-01 | Add DocumentName and DocumentPath display fields to Connection Stat... | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | S | — | [ ] |
| 226 | TASK-002-04-02 | Implement DocumentName and DocumentPath properties with change noti... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | — | [ ] |
| 227 | TASK-002-04-03 | Populate document info from ActiveDocument during connection | `src/AutoCAD/AutoCADService.cs` | S | — | [ ] |
| 228 | TASK-002-04-04 | Implement RefreshStatusCommand to update document info on demand | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | TASK-002-04-02 | [ ] |
| 229 | TASK-002-04-05 | Add "no document open" state handling with control disabling | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | TASK-002-04-02 | [ ] |
| 230 | TASK-002-05-01 | Add PR Number, Project Name, and Job Working Plan Name display fields | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | S | — | [ ] |
| 231 | TASK-002-05-02 | Implement PRNumber, ProjectName, JobWorkingPlanName, and ProjectId ... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | — | [ ] |
| 232 | TASK-002-05-03 | Implement process_pr_no equivalent in AutoCADService | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 233 | TASK-002-05-04 | Integrate IOdooService.GetProjectAsync() lookup for PR-to-ProjectId... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | TASK-002-05-02 | [ ] |
| 234 | TASK-002-05-05 | Add warning display when PR number is not found in Odoo projects | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | TASK-002-05-04 | [ ] |
| 235 | TASK-002-06-01 | Add "Clear IDs (This Layout)" and "Clear All" buttons to Actions Bar | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | S | — | [ ] |
| 236 | TASK-002-06-02 | Implement ClearTableIdsCommand (single layout) with confirmation di... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | TASK-002-06-07, TASK-002-06-08 | [ ] |
| 237 | TASK-002-06-03 | Implement ClearAllTableIdsCommand (all layouts) with confirmation d... | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | TASK-002-06-07, TASK-002-06-08 | [ ] |
| 238 | TASK-002-06-04 | Implement clear_table_id equivalent in AutoCADService for single la... | `src/AutoCAD/AutoCADService.cs` | M | TASK-002-06-06 | [ ] |
| 239 | TASK-002-06-05 | Implement clear_all_tables_id equivalent for all layouts | `src/AutoCAD/AutoCADService.cs` | M | TASK-002-06-04 | [ ] |
| 240 | TASK-002-06-06 | Implement set_attribute_value equivalent for writing block attributes | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 241 | TASK-002-06-07 | Define SetTableValue() and ClearSelection() in IAutoCADService inte... | `src/AutoCAD/IAutoCADService.cs` | S | — | [ ] |
| 242 | TASK-002-06-08 | Add success/failure notification display after write-back operations | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | — | [ ] |
| 243 | TASK-002-07-01 | Add connection status indicator and AutoCAD version info to XAML | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | S | — | [ ] |
| 244 | TASK-002-07-02 | Add "Refresh Status" button to the Connection Status panel | `src/OdooAutoCAD.App/Views/Pages/AutoCADPage.xaml` | S | — | [ ] |
| 245 | TASK-002-07-03 | Implement GetStatusAsync() in AutoCADService for connection health ... | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 246 | TASK-002-07-04 | Define GetStatusAsync() returning AutoCADStatus in IAutoCADService | `src/AutoCAD/IAutoCADService.cs` | S | — | [ ] |
| 247 | TASK-002-07-05 | Implement RefreshStatusCommand in ViewModel | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | S | TASK-002-07-03 | [ ] |
| 248 | TASK-002-07-06 | Implement periodic connection health check using DispatcherTimer | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | — | [ ] |
| 249 | TASK-002-07-07 | Implement COM timeout detection and retry logic via IGUIProxy | `src/Threading/IGUIProxy.cs` | M | — | [ ] |
| 250 | TASK-002-07-08 | Handle unexpected disconnection with state update and notification | `src/OdooAutoCAD.App/ViewModels/AutoCADViewModel.cs` | M | TASK-002-07-06 | [ ] |
| 251 | TASK-003-06-01 | Add product search TextBox and button to XAML | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 252 | TASK-003-06-02 | Add ProductSearchTerm property to ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 253 | TASK-003-06-03 | Implement SearchProductsCommand | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-06-02, TASK-003-06-04 | [ ] |
| 254 | TASK-003-06-04 | Implement SearchProductsAsync in OdooService | `src/Odoo/OdooService.cs` | M | — | [ ] |
| 255 | TASK-003-06-05 | Add search input validation and IsEnabled binding | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-06-01 | [ ] |
| 256 | TASK-003-06-06 | Implement clear search to restore cached product list | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | — | [ ] |
| 257 | TASK-003-07-01 | Create Project Search panel in XAML | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | L | — | [ ] |
| 258 | TASK-003-07-02 | Add project search properties to ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 259 | TASK-003-07-03 | Implement SearchProjectsCommand | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-07-02 | [ ] |
| 260 | TASK-003-07-04 | Implement SelectProjectCommand | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-07-02 | [ ] |
| 261 | TASK-003-07-05 | Implement GetProjectAsync for single-project retrieval | `src/Odoo/OdooService.cs` | M | — | [ ] |
| 262 | TASK-003-07-06 | Bind search field IsEnabled to IsConnected | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-07-01 | [ ] |
| 263 | TASK-003-07-07 | Add "Select Project" button below results DataGrid | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | TASK-003-07-01 | [ ] |
| 264 | TASK-003-08-01 | Add last sync timestamp labels to XAML panels | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 265 | TASK-003-08-02 | Add LastSyncTime property to ViewModel | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 266 | TASK-003-08-03 | Implement GetLastSyncTime in OdooService | `src/Odoo/OdooService.cs` | S | — | [ ] |
| 267 | TASK-003-08-04 | Persist last sync time to local database | `src/OdooAutoCAD.Data/Context/AppDbContext.cs` | M | — | [ ] |
| 268 | TASK-003-08-05 | Update LastSyncTime after each successful sync operation | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-08-02 | [ ] |
| 269 | TASK-003-08-06 | Create DateTime-to-string converter with "Never synced" fallback | `...CAD.App/Converters/DateTimeToStringConverter.cs` | M | — | [ ] |
| 270 | TASK-003-09-01 | Add server info labels to Connection Status panel in XAML | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 271 | TASK-003-09-02 | Add ServerVersion, ConnectedDatabase, ConnectedUsername properties | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 272 | TASK-003-09-03 | Implement GetStatusAsync in OdooService | `src/Odoo/OdooService.cs` | M | — | [ ] |
| 273 | TASK-003-09-04 | Call version_info after successful authentication | `src/Odoo/OdooService.cs` | M | TASK-003-09-03 | [ ] |
| 274 | TASK-003-09-05 | Populate server info after ConnectCommand execution | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-09-02 | [ ] |
| 275 | TASK-003-09-06 | Clear server info on DisconnectCommand execution | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | TASK-003-09-05 | [ ] |
| 276 | TASK-003-11-01 | Add category filter ComboBox to XAML | `...AutoCAD.App/Views/Pages/OdooConnectionPage.xaml` | S | — | [ ] |
| 277 | TASK-003-11-02 | Add SelectedCategoryId and Categories properties | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | S | — | [ ] |
| 278 | TASK-003-11-03 | Implement FilterByCategoryCommand | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | TASK-003-11-02 | [ ] |
| 279 | TASK-003-11-04 | Implement GetProductsByCategoryAsync with child_of operator | `src/Odoo/OdooService.cs` | M | — | [ ] |
| 280 | TASK-003-11-05 | Fetch and populate category list from Odoo | `src/Odoo/OdooService.cs` | L | — | [ ] |
| 281 | TASK-003-11-06 | Add category filter validation | `...toCAD.App/ViewModels/OdooConnectionViewModel.cs` | M | — | [ ] |

### Sprint 8: BOQ P2 & PR Feature Completion
> **Focus**: BOQ product mapping/validation errors/clear IDs/manage mappings/progress, PR details/status/filter/feedback/totals
> **User Stories**: US-004-08, US-004-09, US-004-10, US-004-11, US-004-12, US-005-03, US-005-05, US-005-06, US-005-07, US-005-08
> **Tasks**: 69 (52S + 16M + 1L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 282 | TASK-004-08-01 | Create ProductMappingDialog XAML with AutoCAD name display, search ... | `...CAD.App/Views/Dialogs/ProductMappingDialog.xaml` | M | — | [ ] |
| 283 | TASK-004-08-02 | Implement ProductMappingDialogViewModel with search command and Odo... | `...Models/Dialogs/ProductMappingDialogViewModel.cs` | M | — | [ ] |
| 284 | TASK-004-08-03 | Implement GetProductMappings() in IBOQProcessor returning case-inse... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 285 | TASK-004-08-04 | Implement SetProductMapping() in IBOQProcessor for persisting user-... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 286 | TASK-004-08-05 | Implement SearchProductsAsync() in IOdooService for querying Odoo p... | `src/Odoo/OdooService.cs` | M | — | [ ] |
| 287 | TASK-004-08-06 | Wire "Map Product" action button in validation results panel to ope... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-08-01, TASK-004-08-02 | [ ] |
| 288 | TASK-004-08-07 | Update BOQEntryRow mapping state after successful product mapping | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-08-06 | [ ] |
| 289 | TASK-004-08-08 | Trigger re-validation of affected entries after mapping is applied | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-08-07 | [ ] |
| 290 | TASK-004-09-01 | Create validation results panel XAML as a ListView of BOQValidation... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | M | — | [ ] |
| 291 | TASK-004-09-02 | Implement DataTemplate for validation error items with severity ico... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | M | TASK-004-09-01 | [ ] |
| 292 | TASK-004-09-03 | Implement NavigateToErrorCommand that selects/scrolls to the corres... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | — | [ ] |
| 293 | TASK-004-09-04 | Populate ValidationErrors ObservableCollection from IBOQProcessor.V... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |
| 294 | TASK-004-09-05 | Implement "Map Product" inline action that opens ProductMappingDial... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |
| 295 | TASK-004-09-06 | Implement "Ignore" inline action that dismisses Warning-severity en... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |
| 296 | TASK-004-09-07 | Add error/warning count header to the validation results panel | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | TASK-004-09-01 | [ ] |
| 297 | TASK-004-10-01 | Add "Clear All IDs" button to extraction panel in BOQ Manager page | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 298 | TASK-004-10-02 | Implement ClearAllIdsCommand that calls IGUIProxy for COM clear ope... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | TASK-004-10-05 | [ ] |
| 299 | TASK-004-10-03 | Implement clear_all_tables_id equivalent in IAutoCADService | `src/AutoCAD/AutoCADService.cs` | M | — | [ ] |
| 300 | TASK-004-10-04 | Register clear handler in IGUIProxy for STA thread execution | `src/OdooAutoCAD.App/App.xaml.cs` | S | TASK-004-10-03 | [ ] |
| 301 | TASK-004-10-05 | Add confirmation dialog before executing clear operation | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | — | [ ] |
| 302 | TASK-004-10-06 | Update DataGrid DetailId values to empty after clear completes | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-10-02 | [ ] |
| 303 | TASK-004-11-01 | Add "Manage Mappings" button to BOQ Manager page that opens mapping... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 304 | TASK-004-11-02 | Create mapping management section or dialog with DataGrid showing a... | `...CAD.App/Views/Dialogs/ProductMappingDialog.xaml` | M | — | [ ] |
| 305 | TASK-004-11-03 | Implement mapping management ViewModel with Add, Delete, and Modify... | `...Models/Dialogs/ProductMappingDialogViewModel.cs` | L | — | [ ] |
| 306 | TASK-004-11-04 | Bind ProductMappings ObservableCollection to the management view Da... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | S | TASK-004-11-02, TASK-004-11-03 | [ ] |
| 307 | TASK-004-11-05 | Implement GetProductMappings() to return all current mappings for d... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 308 | TASK-004-11-06 | Implement Add/Remove mapping operations through IBOQProcessor inter... | `src/BOQ/BOQProcessor.cs` | S | — | [ ] |
| 309 | TASK-004-11-07 | Wire Odoo product search for adding/modifying mappings | `...Models/Dialogs/ProductMappingDialogViewModel.cs` | S | — | [ ] |
| 310 | TASK-004-12-01 | Add ProgressBar XAML bound to ExtractionProgressPercent with BoolTo... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 311 | TASK-004-12-02 | Add progress text label bound to ExtractionProgress | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 312 | TASK-004-12-03 | Implement IProgress<T> reporting in ExtractCommand to update progre... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | — | [ ] |
| 313 | TASK-004-12-04 | Add push progress bar and text bound to IsPushing and push progress... | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 314 | TASK-004-12-05 | Implement IProgress<T> reporting in PushToOdooCommand to update pro... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | — | [ ] |
| 315 | TASK-004-12-06 | Implement Cancel button with CancellationTokenSource for aborting l... | `...ooAutoCAD.App/ViewModels/BOQManagerViewModel.cs` | M | — | [ ] |
| 316 | TASK-004-12-07 | Add busy state styling for Extract and Push buttons during operations | `...OdooAutoCAD.App/Views/Pages/BOQManagerPage.xaml` | S | — | [ ] |
| 317 | TASK-005-03-01 | Create PR detail panel layout with header info | `Views/Pages/PurchaseRequisitionPage.xaml` | M | — | [ ] |
| 318 | TASK-005-03-02 | Create PR lines DataGrid | `Views/Pages/PurchaseRequisitionPage.xaml` | M | TASK-005-03-01 | [ ] |
| 319 | TASK-005-03-03 | Implement SelectPRCommand to load line items | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 320 | TASK-005-03-04 | Map PRLine to PRLineDisplayItem with computed LineTotal | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-03-03 | [ ] |
| 321 | TASK-005-03-05 | Implement SelectedPRTotal computed property | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-03-03 | [ ] |
| 322 | TASK-005-03-06 | Add quantity validation highlighting | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-03-02 | [ ] |
| 323 | TASK-005-03-07 | Add placeholder state for empty detail panel | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-03-01 | [ ] |
| 324 | TASK-005-05-01 | Create reusable state badge XAML template | `Views/Pages/PurchaseRequisitionPage.xaml` | M | — | [ ] |
| 325 | TASK-005-05-02 | Implement StateMap dictionary for state-to-display mapping | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 326 | TASK-005-05-03 | Compute badge properties on PRDisplayItem | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-05-02 | [ ] |
| 327 | TASK-005-05-04 | Apply state badge to PR list ListView item template | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-05-01 | [ ] |
| 328 | TASK-005-05-05 | Apply state badge to PR detail header | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-05-01 | [ ] |
| 329 | TASK-005-05-06 | Ensure real-time state change propagation | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-05-03 | [ ] |
| 330 | TASK-005-06-01 | Add filter dropdown ComboBox | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 331 | TASK-005-06-02 | Add sort order dropdown ComboBox | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 332 | TASK-005-06-03 | Implement SelectedStateFilter property | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-06-01 | [ ] |
| 333 | TASK-005-06-04 | Implement SelectedSortOrder property | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-06-02 | [ ] |
| 334 | TASK-005-06-05 | Implement FilteredPRs computed collection | `ViewModels/PurchaseRequisitionViewModel.cs` | M | TASK-005-06-03, TASK-005-06-04 | [ ] |
| 335 | TASK-005-06-06 | Implement FilterByStateCommand and SortByCommand | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-06-05 | [ ] |
| 336 | TASK-005-06-07 | Validate filter state values | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 337 | TASK-005-07-01 | Add status message area in conversion controls section | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 338 | TASK-005-07-02 | Implement success feedback logic | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 339 | TASK-005-07-03 | Implement failure feedback logic | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 340 | TASK-005-07-04 | Implement empty result feedback | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 341 | TASK-005-07-05 | Implement network timeout detection | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 342 | TASK-005-07-06 | Add LastConversionTime display in project context panel | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 343 | TASK-005-07-07 | Style status message area with visual differentiation | `Views/Pages/PurchaseRequisitionPage.xaml` | S | TASK-005-07-01 | [ ] |
| 344 | TASK-005-08-01 | Create summary bar layout at page bottom | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 345 | TASK-005-08-02 | Implement state count properties | `ViewModels/PurchaseRequisitionViewModel.cs` | S | — | [ ] |
| 346 | TASK-005-08-03 | Implement UpdateSummaryCounts method | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-08-02 | [ ] |
| 347 | TASK-005-08-04 | Call UpdateSummaryCounts after state-changing operations | `ViewModels/PurchaseRequisitionViewModel.cs` | S | TASK-005-08-03 | [ ] |
| 348 | TASK-005-08-05 | Display TotalCount in project context panel | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 349 | TASK-005-08-06 | Ensure LineCount displayed in PR list item template | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |
| 350 | TASK-005-08-07 | Bind SelectedPRTotal in detail view with currency formatting | `Views/Pages/PurchaseRequisitionPage.xaml` | S | — | [ ] |

### Sprint 9: Settings Advanced & Keyboard Shortcuts
> **Focus**: AutoCAD timeout, MCP port/auto-start, theme/language/log level, cache/export/import config, app info, keyboard shortcuts
> **User Stories**: US-007-04, US-007-05, US-007-06, US-007-07, US-007-08, US-007-09, US-007-10, US-007-11, US-007-12, US-007-13, US-008-07
> **Tasks**: 71 (55S + 15M + 1L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 351 | TASK-007-04-01 | Add AutoCAD tab with ProgId, Connection Timeout, and Retry Attempts... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 352 | TASK-007-04-02 | Add AutoCADProgId, AutoCADConnectionTimeout, and AutoCADRetryAttemp... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 353 | TASK-007-04-03 | Implement validation for timeout (5-120) and retry attempts (1-10) ... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-04-02 | [ ] |
| 354 | TASK-007-04-04 | Add numeric input constraints to XAML fields (integer-only input) | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 355 | TASK-007-04-05 | Persist AutoCAD settings to ServerConfigs table on Save | `ViewModels/SettingsViewModel.cs` | S | TASK-007-04-02 | [ ] |
| 356 | TASK-007-04-06 | Load AutoCAD settings from ServerConfigs table on page initializati... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-04-05 | [ ] |
| 357 | TASK-007-05-01 | Add MCP tab with SSE Server Port input field in XAML | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 358 | TASK-007-05-02 | Add MCPPort property to ViewModel with default value 8084 | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 359 | TASK-007-05-03 | Implement validation for port range (1024-65535) with inline error ... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-05-02 | [ ] |
| 360 | TASK-007-05-04 | Implement optional port-in-use detection and display non-blocking w... | `ViewModels/SettingsViewModel.cs` | M | TASK-007-05-02 | [ ] |
| 361 | TASK-007-05-05 | Persist MCP port to ServerConfigs table and appsettings.json on Save | `ViewModels/SettingsViewModel.cs` | S | TASK-007-05-02 | [ ] |
| 362 | TASK-007-05-06 | Load MCP port from ServerConfigs table on page initialization with ... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-05-05 | [ ] |
| 363 | TASK-007-06-01 | Add MCP auto-start toggle/checkbox to MCP tab in XAML | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 364 | TASK-007-06-02 | Add MCPAutoStart boolean property to ViewModel with default value f... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 365 | TASK-007-06-03 | Persist MCP auto-start to ServerConfigs table and appsettings.json ... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-06-02 | [ ] |
| 366 | TASK-007-06-04 | Load MCP auto-start from ServerConfigs table on page initialization... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-06-02, TASK-007-06-03 | [ ] |
| 367 | TASK-007-06-05 | Integrate auto-start check in application startup logic | `` | M | TASK-007-06-03 | [ ] |
| 368 | TASK-007-07-01 | Add theme radio buttons (System, Light, Dark) to Appearance tab in ... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 369 | TASK-007-07-02 | Add SelectedTheme property to ViewModel that triggers live preview ... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 370 | TASK-007-07-03 | Implement IThemeService with ApplyTheme() to swap WPF ResourceDicti... | `...IThemeService.cs` and `Services/ThemeService.cs` | M | TASK-007-07-04 | [ ] |
| 371 | TASK-007-07-04 | Create Light and Dark theme ResourceDictionary XAML files | `...heme.xaml` and `Resources/Themes/DarkTheme.xaml` | M | — | [ ] |
| 372 | TASK-007-07-05 | Implement "System" theme detection following Windows OS light/dark ... | `Services/ThemeService.cs` | M | TASK-007-07-03 | [ ] |
| 373 | TASK-007-07-06 | Persist theme selection to UserPreferences table on Save | `ViewModels/SettingsViewModel.cs` | S | TASK-007-07-02 | [ ] |
| 374 | TASK-007-07-07 | Load theme preference on application startup and apply before main ... | `App.xaml.cs` | S | TASK-007-07-06 | [ ] |
| 375 | TASK-007-08-01 | Add language selector dropdown to Appearance tab in XAML | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 376 | TASK-007-08-02 | Add restart notification text that appears when language differs fr... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 377 | TASK-007-08-03 | Add SelectedLanguage property to ViewModel with supported locale codes | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 378 | TASK-007-08-04 | Persist language selection to UserPreferences table on Save | `ViewModels/SettingsViewModel.cs` | S | TASK-007-08-03 | [ ] |
| 379 | TASK-007-08-05 | Load language preference on application startup and set CultureInfo... | `App.xaml.cs` | M | TASK-007-08-04 | [ ] |
| 380 | TASK-007-08-06 | Create resource files (.resx) for English and Traditional Chinese s... | `...resx` and `Resources/Strings/Strings.zh-TW.resx` | L | TASK-007-08-03 | [ ] |
| 381 | TASK-007-09-01 | Add Logging section to Advanced tab with log level dropdown, log pa... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 382 | TASK-007-09-02 | Add SelectedLogLevel and LogFilePath properties to ViewModel | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 383 | TASK-007-09-03 | Implement OpenLogFolderCommand that opens log directory in Windows ... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 384 | TASK-007-09-04 | Implement validation for log level (must be one of the six valid op... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-09-02 | [ ] |
| 385 | TASK-007-09-05 | Persist log level to UserPreferences table on Save | `ViewModels/SettingsViewModel.cs` | S | TASK-007-09-02 | [ ] |
| 386 | TASK-007-09-06 | Apply log level change at runtime by reconfiguring the logging prov... | `ViewModels/SettingsViewModel.cs` | M | TASK-007-09-02 | [ ] |
| 387 | TASK-007-09-07 | Load log level from UserPreferences table on page initialization wi... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-09-05 | [ ] |
| 388 | TASK-007-10-01 | Add "Clear Cache" button to Data Management section of Advanced tab... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 389 | TASK-007-10-02 | Implement ClearCacheCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M | — | [ ] |
| 390 | TASK-007-10-03 | Query BOQCache and ProductMapping record counts for confirmation di... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 391 | TASK-007-10-04 | Implement confirmation dialog with record counts and Cancel/Continu... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-10-02, TASK-007-10-03 | [ ] |
| 392 | TASK-007-10-05 | Execute RemoveRange on BOQCache and ProductMapping tables via AppDb... | `ViewModels/SettingsViewModel.cs` | M | TASK-007-10-04 | [ ] |
| 393 | TASK-007-10-06 | Invalidate in-memory caches in IOdooService after database clear | `` | S | TASK-007-10-05 | [ ] |
| 394 | TASK-007-10-07 | Add success/error result feedback after cache clear operation | `ViewModels/SettingsViewModel.cs` | S | TASK-007-10-02, TASK-007-10-05 | [ ] |
| 395 | TASK-007-10-08 | Log cache clear operation to SyncLog table | `ViewModels/SettingsViewModel.cs` | S | TASK-007-10-02 | [ ] |
| 396 | TASK-007-11-01 | Add "Export Configuration" button to Data Management section of Adv... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 397 | TASK-007-11-02 | Implement ExportConfigCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M | — | [ ] |
| 398 | TASK-007-11-03 | Implement IFileDialogService.ShowSaveFileDialog() for JSON file sel... | `...gService.cs` and `Services/FileDialogService.cs` | S | — | [ ] |
| 399 | TASK-007-11-04 | Serialize current settings to JSON excluding sensitive fields (token) | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M | — | [ ] |
| 400 | TASK-007-11-05 | Write serialized JSON to the user-selected file path with error han... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-11-02, TASK-007-11-04 | [ ] |
| 401 | TASK-007-11-06 | Display success or error feedback after export operation | `ViewModels/SettingsViewModel.cs` | S | TASK-007-11-05 | [ ] |
| 402 | TASK-007-11-07 | Log export operation to SyncLog table | `ViewModels/SettingsViewModel.cs` | S | TASK-007-11-02 | [ ] |
| 403 | TASK-007-12-01 | Add "Import Configuration" button to Data Management section of Adv... | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 404 | TASK-007-12-02 | Implement ImportConfigCommand as IAsyncRelayCommand in ViewModel | `ViewModels/SettingsViewModel.cs` | M | — | [ ] |
| 405 | TASK-007-12-03 | Implement IFileDialogService.ShowOpenFileDialog() for JSON file sel... | `...gService.cs` and `Services/FileDialogService.cs` | S | — | [ ] |
| 406 | TASK-007-12-04 | Implement JSON parsing and validation for structure and expected to... | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | M | — | [ ] |
| 407 | TASK-007-12-05 | Map imported JSON values to ViewModel properties and set HasUnsaved... | `ViewModels/SettingsViewModel.cs` | M | TASK-007-12-02, TASK-007-12-04 | [ ] |
| 408 | TASK-007-12-06 | Display success, partial-import warning, or error feedback based on... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-12-04 | [ ] |
| 409 | TASK-007-12-07 | Log import operation to SyncLog table | `ViewModels/SettingsViewModel.cs` | S | TASK-007-12-02 | [ ] |
| 410 | TASK-007-13-01 | Add Application Info section to Advanced tab with read-only fields | `Views/Pages/SettingsPage.xaml` | S | — | [ ] |
| 411 | TASK-007-13-02 | Add AppVersion, DatabasePath, ConfigFilePath, and DotNetRuntime rea... | `ViewModels/SettingsViewModel.cs` | S | — | [ ] |
| 412 | TASK-007-13-03 | Populate AppVersion from assembly version or application metadata | `ViewModels/SettingsViewModel.cs` | S | TASK-007-13-02 | [ ] |
| 413 | TASK-007-13-04 | Populate DatabasePath from AppDbContext connection string | `ViewModels/SettingsViewModel.cs` | S | TASK-007-13-02 | [ ] |
| 414 | TASK-007-13-05 | Populate ConfigFilePath from ConfigurationLoader or AppDomain.Curre... | `ViewModels/SettingsViewModel.cs` | S | TASK-007-13-02 | [ ] |
| 415 | TASK-007-13-06 | Populate DotNetRuntime from RuntimeInformation.FrameworkDescription | `ViewModels/SettingsViewModel.cs` | S | TASK-007-13-02 | [ ] |
| 416 | TASK-007-13-07 | Style Application Info fields as read-only with distinct visual tre... | `Views/Pages/SettingsPage.xaml` | S | TASK-007-13-01 | [ ] |
| 417 | TASK-008-07-01 | Define InputBindings for Ctrl+1 through Ctrl+7 navigation shortcuts | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | M | — | [ ] |
| 418 | TASK-008-07-02 | Define InputBinding for F5 Refresh shortcut | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | — | [ ] |
| 419 | TASK-008-07-03 | Define InputBinding for Alt+Left GoBack shortcut | `src/OdooAutoCAD.App/Views/MainWindow.xaml` | S | TASK-008-07-04 | [ ] |
| 420 | TASK-008-07-04 | Add GoBackCommand to MainViewModel | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | — | [ ] |
| 421 | TASK-008-07-05 | Ensure NavigateCommand accepts string parameter for keyboard shortc... | `src/OdooAutoCAD.App/ViewModels/MainViewModel.cs` | S | — | [ ] |

---

## Phase 4: AI Integration

### Sprint 10: MCP Server Foundation
> **Focus**: MCP server start/stop, server status, test connection, tool registry, configuration, recent activity
> **User Stories**: US-006-01, US-006-02, US-006-03, US-006-04, US-006-05, US-006-06, US-001-03
> **Tasks**: 41 (23S + 16M + 2L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 422 | TASK-006-01-01 | Implement StartServerCommand in MCPViewModel with port validation a... | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 423 | TASK-006-01-02 | Add Start button with toggle binding and disabled state in XAML | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 424 | TASK-006-01-03 | Implement MCPSSEServer.StartAsync() with port availability check an... | `MCP/Server/MCPSSEServer.cs` | L | — | [ ] |
| 425 | TASK-006-01-04 | Register all MCP tools via MCPToolRegistry.RegisterAllTools() durin... | `MCP/Tools/MCPToolRegistry.cs` | L | TASK-006-01-03 | [ ] |
| 426 | TASK-006-01-05 | Initialize IGUIProxy and start DispatcherTimer for proxy polling on... | `Views/Pages/MCPAssistantPage.xaml.cs` | M | TASK-006-01-01 | [ ] |
| 427 | TASK-006-01-06 | Wire status callback to update ViewModel properties on state change | `ViewModels/MCPViewModel.cs` | S | TASK-006-01-01 | [ ] |
| 428 | TASK-006-02-01 | Implement StopServerCommand in MCPViewModel with graceful shutdown ... | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 429 | TASK-006-02-02 | Enhance MCPSSEServer.StopAsync() to close all SSE connections and r... | `MCP/Server/MCPSSEServer.cs` | M | — | [ ] |
| 430 | TASK-006-02-03 | Stop DispatcherTimer and clean up IGUIProxy resources on server stop | `Views/Pages/MCPAssistantPage.xaml.cs` | S | — | [ ] |
| 431 | TASK-006-02-04 | Update XAML button binding to show "Start Server" when stopped and ... | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 432 | TASK-006-02-05 | Log server shutdown event and connection closures in activity log | `ViewModels/MCPViewModel.cs` | S | TASK-006-02-01 | [ ] |
| 433 | TASK-006-03-01 | Create Server Control panel XAML layout with status indicator, port... | `Views/Pages/MCPAssistantPage.xaml` | M | — | [ ] |
| 434 | TASK-006-03-02 | Implement BoolToColorConverter for mapping IsServerRunning to green... | `Converters/BoolToColorConverter.cs` | S | TASK-006-03-01 | [ ] |
| 435 | TASK-006-03-03 | Add ViewModel properties for server status display | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 436 | TASK-006-03-04 | Implement status callback handler that updates ViewModel on server ... | `ViewModels/MCPViewModel.cs` | S | TASK-006-03-03 | [ ] |
| 437 | TASK-006-03-05 | Implement uptime timer that updates ServerUptime every second while... | `ViewModels/MCPViewModel.cs` | S | TASK-006-03-03 | [ ] |
| 438 | TASK-006-03-06 | Expose ActiveConnections count and IsInitialized from MCPSSEServer ... | `MCP/Server/MCPSSEServer.cs` | M | — | [ ] |
| 439 | TASK-006-04-01 | Implement TestConnectionCommand in MCPViewModel that calls health a... | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 440 | TASK-006-04-02 | Add Test Connection button to XAML with CanExecute bound to IsServe... | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 441 | TASK-006-04-03 | Implement two-phase test: HTTP GET /health followed by JSON-RPC tes... | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 442 | TASK-006-04-04 | Display test results in the status panel | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 443 | TASK-006-04-05 | Handle timeout and failure scenarios with appropriate user-facing e... | `ViewModels/MCPViewModel.cs` | S | TASK-006-04-01 | [ ] |
| 444 | TASK-006-05-01 | Create Registered Tools panel XAML with grouped ListView and catego... | `Views/Pages/MCPAssistantPage.xaml` | M | — | [ ] |
| 445 | TASK-006-05-02 | Implement MCPToolViewModel with display properties | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 446 | TASK-006-05-03 | Populate RegisteredTools ObservableCollection from MCPToolRegistry.... | `ViewModels/MCPViewModel.cs` | S | TASK-006-05-02 | [ ] |
| 447 | TASK-006-05-04 | Add GroupStyle to ItemsControl for category-based grouping with Col... | `Views/Pages/MCPAssistantPage.xaml` | M | TASK-006-05-01 | [ ] |
| 448 | TASK-006-05-05 | Display prerequisite badges next to tool names using DataTemplate | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 449 | TASK-006-05-06 | Ensure MCPToolRegistry.RegisterTool() triggers collection update vi... | `MCP/Tools/MCPToolRegistry.cs` | M | — | [ ] |
| 450 | TASK-006-06-01 | Create Configuration panel XAML layout with server settings, endpoi... | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 451 | TASK-006-06-02 | Add ViewModel properties for computed endpoint URLs | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 452 | TASK-006-06-03 | Display GUI Proxy status and pending request count in Configuration... | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 453 | TASK-006-06-04 | Display prerequisite legend in Configuration panel | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 454 | TASK-006-06-05 | Add client configuration guidance text for Gemini CLI and Claude Co... | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 455 | TASK-001-03-01 | Create Current Project info panel in XAML | `Views/Pages/DashboardPage.xaml` | M | — | [ ] |
| 456 | TASK-001-03-02 | Add ViewModel properties for project info and sync time | `ViewModels/DashboardViewModel.cs` | S | — | [ ] |
| 457 | TASK-001-03-03 | Implement logic to fetch current project context from IAutoCADService | `ViewModels/DashboardViewModel.cs` | M | TASK-001-03-02 | [ ] |
| 458 | TASK-001-03-04 | Implement logic to retrieve LastSyncTime from IOdooService | `ViewModels/DashboardViewModel.cs` | S | TASK-001-03-02 | [ ] |
| 459 | TASK-001-03-05 | Create Recent Operations summary section in XAML | `Views/Pages/DashboardPage.xaml` | M | TASK-001-03-01 | [ ] |
| 460 | TASK-001-03-06 | Add ViewModel properties and logic for recent operation counts | `ViewModels/DashboardViewModel.cs` | M | TASK-001-03-02 | [ ] |
| 461 | TASK-001-03-07 | Handle "no project context" state | `ViewModels/DashboardViewModel.cs` | S | TASK-001-03-02 | [ ] |
| 462 | TASK-001-03-08 | Subscribe to AutoCAD document change events to refresh project info | `ViewModels/DashboardViewModel.cs` | S | TASK-001-03-02, TASK-001-03-03 | [ ] |

### Sprint 11: MCP Monitoring & Configuration
> **Focus**: MCP connection monitoring, activity logs, tool prerequisites, restart, tool execution, port config, MCP dashboard status
> **User Stories**: US-006-07, US-006-08, US-006-09, US-006-10, US-006-11, US-006-12, US-001-04
> **Tasks**: 38 (24S + 12M + 2L)

| # | Task ID | Title | Target | Est | Depends On | Status |
|---|---------|-------|--------|-----|------------|--------|
| 463 | TASK-006-07-01 | Expose ActiveConnections count property from MCPSSEServer with chan... | `MCP/Server/MCPSSEServer.cs` | S | — | [ ] |
| 464 | TASK-006-07-02 | Bind ActiveConnectionCount ViewModel property to MCPSSEServer.Activ... | `ViewModels/MCPViewModel.cs` | S | TASK-006-07-01 | [ ] |
| 465 | TASK-006-07-03 | Display connection count in the Server Control panel XAML | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 466 | TASK-006-07-04 | Update connection count on SSE connect/disconnect events using the ... | `MCP/Server/MCPSSEServer.cs` | S | — | [ ] |
| 467 | TASK-006-08-01 | Create Activity Log panel XAML with scrollable ListBox, Clear Log b... | `Views/Pages/MCPAssistantPage.xaml` | M | — | [ ] |
| 468 | TASK-006-08-02 | Implement LogEntry model with Timestamp, Level, Message, and Source... | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 469 | TASK-006-08-03 | Add ActivityLog ObservableCollection and log management methods to ... | `ViewModels/MCPViewModel.cs` | M | TASK-006-08-02 | [ ] |
| 470 | TASK-006-08-04 | Implement auto-scroll behavior with IsAutoScrollEnabled toggle and ... | `Views/Pages/MCPAssistantPage.xaml.cs` | S | TASK-006-08-01 | [ ] |
| 471 | TASK-006-08-05 | Implement ClearLogCommand and ToggleAutoScrollCommand relay commands | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 472 | TASK-006-08-06 | Wire MCPSSEServer events to the ActivityLog collection | `ViewModels/MCPViewModel.cs` | M | TASK-006-08-03 | [ ] |
| 473 | TASK-006-09-01 | Add RequiresAutoCAD and RequiresOdoo boolean properties to MCPToolV... | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 474 | TASK-006-09-02 | Display [A] and [O] badges with warning indicators for unmet prereq... | `Views/Pages/MCPAssistantPage.xaml` | S | TASK-006-09-01 | [ ] |
| 475 | TASK-006-09-03 | Add prerequisite legend section to the Configuration panel XAML | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 476 | TASK-006-09-04 | Implement prerequisite checking in MCPToolRegistry.ExecuteToolAsync... | `MCP/Tools/MCPToolRegistry.cs` | M | — | [ ] |
| 477 | TASK-006-09-05 | Query IAutoCADService.IsConnected and IOdooService.IsConnected for ... | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 478 | TASK-006-10-01 | Implement RestartServerCommand in MCPViewModel that chains StopAsyn... | `ViewModels/MCPViewModel.cs` | M | — | [ ] |
| 479 | TASK-006-10-02 | Add Restart button to Server Control panel XAML | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 480 | TASK-006-10-03 | Implement MCPSSEServer.RestartAsync() with 1-second delay between s... | `MCP/Server/MCPSSEServer.cs` | M | — | [ ] |
| 481 | TASK-006-10-04 | Update ServerStatus property through intermediate states during res... | `ViewModels/MCPViewModel.cs` | S | TASK-006-10-01 | [ ] |
| 482 | TASK-006-10-05 | Log restart operation start and completion in activity log | `ViewModels/MCPViewModel.cs` | S | TASK-006-10-01 | [ ] |
| 483 | TASK-006-11-01 | Add execution logging hooks in MCPToolRegistry.ExecuteToolAsync() | `MCP/Tools/MCPToolRegistry.cs` | M | — | [ ] |
| 484 | TASK-006-11-02 | Implement tool execution event handler in MCPViewModel to append Lo... | `ViewModels/MCPViewModel.cs` | S | TASK-006-11-01 | [ ] |
| 485 | TASK-006-11-03 | Add Stopwatch-based duration measurement in ExecuteToolAsync | `MCP/Tools/MCPToolRegistry.cs` | S | — | [ ] |
| 486 | TASK-006-11-04 | Log GUI Proxy action queuing and completion events | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 487 | TASK-006-11-05 | Ensure log entries are dispatched to the UI thread for ObservableCo... | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 488 | TASK-006-12-01 | Add editable port TextBox to Configuration panel XAML | `Views/Pages/MCPAssistantPage.xaml` | S | — | [ ] |
| 489 | TASK-006-12-02 | Implement port validation in MCPViewModel | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 490 | TASK-006-12-03 | Update computed endpoint URL properties when ServerPort changes | `ViewModels/MCPViewModel.cs` | S | — | [ ] |
| 491 | TASK-006-12-04 | Persist port configuration to application settings | `ViewModels/MCPViewModel.cs` | L | — | [ ] |
| 492 | TASK-006-12-05 | Add port availability check on server start and display error if po... | `MCP/Server/MCPSSEServer.cs` | M | — | [ ] |
| 493 | TASK-006-12-06 | Display validation error styling for invalid port input | `Views/Pages/MCPAssistantPage.xaml` | S | TASK-006-12-01 | [ ] |
| 494 | TASK-001-04-01 | Create MCP Server status card in XAML | `Views/Pages/DashboardPage.xaml` | M | — | [ ] |
| 495 | TASK-001-04-02 | Add ViewModel properties for MCP status | `ViewModels/DashboardViewModel.cs` | S | — | [ ] |
| 496 | TASK-001-04-03 | Implement ToggleMCPCommand with async start/stop | `ViewModels/DashboardViewModel.cs` | M | TASK-001-04-02 | [ ] |
| 497 | TASK-001-04-04 | Implement async monitoring of MCP SSE server state | `ViewModels/DashboardViewModel.cs` | M | TASK-001-04-02 | [ ] |
| 498 | TASK-001-04-05 | Add error handling for MCP server start failure | `ViewModels/DashboardViewModel.cs` | S | TASK-001-04-03 | [ ] |
| 499 | TASK-001-04-06 | Bind toggle button content and status indicator in XAML | `Views/Pages/DashboardPage.xaml` | S | TASK-001-04-01 | [ ] |
| 500 | TASK-001-04-07 | Integrate with MCP SSE Manager service for server lifecycle management | `ViewModels/DashboardViewModel.cs` | L | TASK-001-04-02, TASK-001-04-03, TASK-001-04-04 | [ ] |

---

## Cross-Sprint Dependencies

| Dependency | Description |
|-----------|-------------|
| Sprint 1 → Sprint 2 | App shell, sidebar, navigation needed for status bar, log panel, dashboard |
| Sprint 1 → Sprint 3 | Window framework needed for AutoCAD/Odoo connection pages |
| Sprint 2 → Sprint 4 | Dashboard connection status depends on Odoo service infrastructure |
| Sprint 3 → Sprint 4 | AutoCAD/Odoo service interfaces needed for BOQ and sync operations |
| Sprint 3 → Sprint 5 | AutoCAD parameter extraction needed for BOQ processing |
| Sprint 4 → Sprint 5 | BOQ extract/review foundation needed for validate/push pipeline |
| Sprint 4 → Sprint 6 | Odoo sync needed for PR conversion; settings depend on config infrastructure |
| Sprint 5 → Sprint 6 | BOQ push to Odoo needed before PR conversion |
| Sprint 3 → Sprint 7 | Core AutoCAD/Odoo connections needed for P2 features |
| Sprint 5 → Sprint 8 | BOQ core pipeline needed for mapping/progress P2 features |
| Sprint 6 → Sprint 8 | PR list/submit needed for PR details/status/filter features |
| Sprint 6 → Sprint 9 | Settings core needed for advanced settings |
| Sprint 1 → Sprint 10 | App shell needed for MCP assistant page |
| Sprint 10 → Sprint 11 | MCP server foundation needed for monitoring/config features |

## Effort Summary by Sprint

| Sprint | Name | S | M | L | Total |
|--------|------|---|---|---|-------|
| 1 | App Shell & Navigation | 22 | 9 | 1 | 32 |
| 2 | Status Bar, Log Panel & Dashboard | 15 | 15 | 1 | 31 |
| 3 | AutoCAD & Odoo Core Connections | 20 | 15 | 5 | 40 |
| 4 | Odoo Advanced & BOQ Foundation | 26 | 10 | 6 | 42 |
| 5 | BOQ Processing Pipeline | 20 | 9 | 4 | 33 |
| 6 | Purchase Requisition & Settings Core | 25 | 19 | 2 | 46 |
| 7 | AutoCAD P2 & Odoo Search/Filter | 31 | 24 | 2 | 57 |
| 8 | BOQ P2 & PR Feature Completion | 52 | 16 | 1 | 69 |
| 9 | Settings Advanced & Keyboard Shortcuts | 55 | 15 | 1 | 71 |
| 10 | MCP Server Foundation | 23 | 16 | 2 | 41 |
| 11 | MCP Monitoring & Configuration | 24 | 12 | 2 | 38 |
| **Total** | | **313** | **160** | **27** | **500** |

## User Story Index

| US ID | Title | Priority | Tasks | Sprint | TASK File |
|-------|-------|----------|-------|--------|-----------|
| US-001-01 | View Connection Status | P1 | 6 | 2 | [TASKS-US-001-01.md](FR-001-dashboard/TASKS-US-001-01.md) |
| US-001-02 | Quick Connect Buttons | P1 | 8 | 2 | [TASKS-US-001-02.md](FR-001-dashboard/TASKS-US-001-02.md) |
| US-001-03 | Recent Activity | P2 | 8 | 10 | [TASKS-US-001-03.md](FR-001-dashboard/TASKS-US-001-03.md) |
| US-001-04 | MCP Server Status | P2 | 7 | 11 | [TASKS-US-001-04.md](FR-001-dashboard/TASKS-US-001-04.md) |
| US-002-01 | Connect to AutoCAD | P1 | 6 | 3 | [TASKS-US-002-01.md](FR-002-autocad-connection/TASKS-US-002-01.md) |
| US-002-02 | View Layouts | P1 | 7 | 3 | [TASKS-US-002-02.md](FR-002-autocad-connection/TASKS-US-002-02.md) |
| US-002-03 | Extract Parameters | P1 | 9 | 3 | [TASKS-US-002-03.md](FR-002-autocad-connection/TASKS-US-002-03.md) |
| US-002-04 | Drawing Info | P1 | 5 | 7 | [TASKS-US-002-04.md](FR-002-autocad-connection/TASKS-US-002-04.md) |
| US-002-05 | PR Project Info | P1 | 5 | 7 | [TASKS-US-002-05.md](FR-002-autocad-connection/TASKS-US-002-05.md) |
| US-002-06 | Clear Table IDs | P1 | 8 | 7 | [TASKS-US-002-06.md](FR-002-autocad-connection/TASKS-US-002-06.md) |
| US-002-07 | Monitor COM Status | P2 | 8 | 7 | [TASKS-US-002-07.md](FR-002-autocad-connection/TASKS-US-002-07.md) |
| US-003-01 | Enter Odoo Credentials | P1 | 6 | 3 | [TASKS-US-003-01.md](FR-003-odoo-connection/TASKS-US-003-01.md) |
| US-003-02 | Test Connection | P1 | 6 | 3 | [TASKS-US-003-02.md](FR-003-odoo-connection/TASKS-US-003-02.md) |
| US-003-03 | View Connection Status | P1 | 6 | 4 | [TASKS-US-003-03.md](FR-003-odoo-connection/TASKS-US-003-03.md) |
| US-003-04 | Disconnect from Odoo | P1 | 6 | 4 | [TASKS-US-003-04.md](FR-003-odoo-connection/TASKS-US-003-04.md) |
| US-003-05 | Sync Product Catalog | P1 | 7 | 4 | [TASKS-US-003-05.md](FR-003-odoo-connection/TASKS-US-003-05.md) |
| US-003-06 | Search Products | P1 | 6 | 7 | [TASKS-US-003-06.md](FR-003-odoo-connection/TASKS-US-003-06.md) |
| US-003-07 | Search Projects | P1 | 7 | 7 | [TASKS-US-003-07.md](FR-003-odoo-connection/TASKS-US-003-07.md) |
| US-003-08 | View Last Sync Time | P1 | 6 | 7 | [TASKS-US-003-08.md](FR-003-odoo-connection/TASKS-US-003-08.md) |
| US-003-09 | View Server Info | P2 | 6 | 7 | [TASKS-US-003-09.md](FR-003-odoo-connection/TASKS-US-003-09.md) |
| US-003-10 | Configure via appsettings.json | P2 | 6 | 3 | [TASKS-US-003-10.md](FR-003-odoo-connection/TASKS-US-003-10.md) |
| US-003-11 | Filter by Category | P2 | 6 | 7 | [TASKS-US-003-11.md](FR-003-odoo-connection/TASKS-US-003-11.md) |
| US-003-12 | Persist Credentials | P2 | 8 | 4 | [TASKS-US-003-12.md](FR-003-odoo-connection/TASKS-US-003-12.md) |
| US-004-01 | Extract BOQ Data | P1 | 8 | 4 | [TASKS-US-004-01.md](FR-004-boq-manager/TASKS-US-004-01.md) |
| US-004-02 | Review Extracted Data | P1 | 7 | 4 | [TASKS-US-004-02.md](FR-004-boq-manager/TASKS-US-004-02.md) |
| US-004-03 | Validate Against Products | P1 | 8 | 5 | [TASKS-US-004-03.md](FR-004-boq-manager/TASKS-US-004-03.md) |
| US-004-04 | Push to Odoo | P1 | 8 | 5 | [TASKS-US-004-04.md](FR-004-boq-manager/TASKS-US-004-04.md) |
| US-004-05 | ID Writeback | P1 | 6 | 5 | [TASKS-US-004-05.md](FR-004-boq-manager/TASKS-US-004-05.md) |
| US-004-06 | View Skipped Rows | P2 | 5 | 5 | [TASKS-US-004-06.md](FR-004-boq-manager/TASKS-US-004-06.md) |
| US-004-07 | View Summary | P2 | 6 | 5 | [TASKS-US-004-07.md](FR-004-boq-manager/TASKS-US-004-07.md) |
| US-004-08 | Map Products | P1 | 8 | 8 | [TASKS-US-004-08.md](FR-004-boq-manager/TASKS-US-004-08.md) |
| US-004-09 | View Validation Errors | P2 | 7 | 8 | [TASKS-US-004-09.md](FR-004-boq-manager/TASKS-US-004-09.md) |
| US-004-10 | Clear IDs | P2 | 6 | 8 | [TASKS-US-004-10.md](FR-004-boq-manager/TASKS-US-004-10.md) |
| US-004-11 | Manage Mappings | P3 | 7 | 8 | [TASKS-US-004-11.md](FR-004-boq-manager/TASKS-US-004-11.md) |
| US-004-12 | Progress Indication | P2 | 7 | 8 | [TASKS-US-004-12.md](FR-004-boq-manager/TASKS-US-004-12.md) |
| US-005-01 | Convert BOQ to PR | P2 | 9 | 6 | [TASKS-US-005-01.md](FR-005-purchase-requisition/TASKS-US-005-01.md) |
| US-005-02 | View PR List | P2 | 9 | 6 | [TASKS-US-005-02.md](FR-005-purchase-requisition/TASKS-US-005-02.md) |
| US-005-03 | View PR Details | P2 | 7 | 8 | [TASKS-US-005-03.md](FR-005-purchase-requisition/TASKS-US-005-03.md) |
| US-005-04 | Submit PR | P2 | 7 | 6 | [TASKS-US-005-04.md](FR-005-purchase-requisition/TASKS-US-005-04.md) |
| US-005-05 | Track PR Status | P2 | 6 | 8 | [TASKS-US-005-05.md](FR-005-purchase-requisition/TASKS-US-005-05.md) |
| US-005-06 | Filter and Sort PRs | P2 | 7 | 8 | [TASKS-US-005-06.md](FR-005-purchase-requisition/TASKS-US-005-06.md) |
| US-005-07 | Conversion Feedback | P2 | 7 | 8 | [TASKS-US-005-07.md](FR-005-purchase-requisition/TASKS-US-005-07.md) |
| US-005-08 | View PR Totals | P2 | 7 | 8 | [TASKS-US-005-08.md](FR-005-purchase-requisition/TASKS-US-005-08.md) |
| US-006-01 | Start MCP Server | P3 | 6 | 10 | [TASKS-US-006-01.md](FR-006-ai-assistant/TASKS-US-006-01.md) |
| US-006-02 | Stop MCP Server | P3 | 5 | 10 | [TASKS-US-006-02.md](FR-006-ai-assistant/TASKS-US-006-02.md) |
| US-006-03 | View Server Status | P3 | 6 | 10 | [TASKS-US-006-03.md](FR-006-ai-assistant/TASKS-US-006-03.md) |
| US-006-04 | Test Connection | P3 | 5 | 10 | [TASKS-US-006-04.md](FR-006-ai-assistant/TASKS-US-006-04.md) |
| US-006-05 | View Tool Registry | P3 | 6 | 10 | [TASKS-US-006-05.md](FR-006-ai-assistant/TASKS-US-006-05.md) |
| US-006-06 | View Configuration | P3 | 5 | 10 | [TASKS-US-006-06.md](FR-006-ai-assistant/TASKS-US-006-06.md) |
| US-006-07 | Monitor Connections | P3 | 4 | 11 | [TASKS-US-006-07.md](FR-006-ai-assistant/TASKS-US-006-07.md) |
| US-006-08 | View Activity Logs | P3 | 6 | 11 | [TASKS-US-006-08.md](FR-006-ai-assistant/TASKS-US-006-08.md) |
| US-006-09 | View Tool Prerequisites | P3 | 5 | 11 | [TASKS-US-006-09.md](FR-006-ai-assistant/TASKS-US-006-09.md) |
| US-006-10 | Restart Server | P3 | 5 | 11 | [TASKS-US-006-10.md](FR-006-ai-assistant/TASKS-US-006-10.md) |
| US-006-11 | View Tool Execution | P3 | 5 | 11 | [TASKS-US-006-11.md](FR-006-ai-assistant/TASKS-US-006-11.md) |
| US-006-12 | Configure Server Port | P3 | 6 | 11 | [TASKS-US-006-12.md](FR-006-ai-assistant/TASKS-US-006-12.md) |
| US-007-01 | Configure Odoo Connection | P2 | 8 | 6 | [TASKS-US-007-01.md](FR-007-settings/TASKS-US-007-01.md) |
| US-007-02 | Switch Environments | P2 | 6 | 6 | [TASKS-US-007-02.md](FR-007-settings/TASKS-US-007-02.md) |
| US-007-03 | Test Odoo Connection | P2 | 7 | 6 | [TASKS-US-007-03.md](FR-007-settings/TASKS-US-007-03.md) |
| US-007-04 | AutoCAD Timeout Settings | P2 | 6 | 9 | [TASKS-US-007-04.md](FR-007-settings/TASKS-US-007-04.md) |
| US-007-05 | MCP Port Configuration | P2 | 6 | 9 | [TASKS-US-007-05.md](FR-007-settings/TASKS-US-007-05.md) |
| US-007-06 | MCP Auto-Start Toggle | P2 | 5 | 9 | [TASKS-US-007-06.md](FR-007-settings/TASKS-US-007-06.md) |
| US-007-07 | Change Theme | P2 | 7 | 9 | [TASKS-US-007-07.md](FR-007-settings/TASKS-US-007-07.md) |
| US-007-08 | Set Language | P2 | 6 | 9 | [TASKS-US-007-08.md](FR-007-settings/TASKS-US-007-08.md) |
| US-007-09 | Change Log Level | P2 | 7 | 9 | [TASKS-US-007-09.md](FR-007-settings/TASKS-US-007-09.md) |
| US-007-10 | Clear Cache | P2 | 8 | 9 | [TASKS-US-007-10.md](FR-007-settings/TASKS-US-007-10.md) |
| US-007-11 | Export Configuration | P2 | 7 | 9 | [TASKS-US-007-11.md](FR-007-settings/TASKS-US-007-11.md) |
| US-007-12 | Import Configuration | P2 | 7 | 9 | [TASKS-US-007-12.md](FR-007-settings/TASKS-US-007-12.md) |
| US-007-13 | View App Info | P2 | 7 | 9 | [TASKS-US-007-13.md](FR-007-settings/TASKS-US-007-13.md) |
| US-008-01 | Sidebar Navigation | P1 | 7 | 1 | [TASKS-US-008-01.md](FR-008-ui-framework/TASKS-US-008-01.md) |
| US-008-02 | Active Page Indicator | P1 | 4 | 1 | [TASKS-US-008-02.md](FR-008-ui-framework/TASKS-US-008-02.md) |
| US-008-03 | Persistent Connection Status | P1 | 5 | 2 | [TASKS-US-008-03.md](FR-008-ui-framework/TASKS-US-008-03.md) |
| US-008-04 | System Log Panel | P1 | 6 | 2 | [TASKS-US-008-04.md](FR-008-ui-framework/TASKS-US-008-04.md) |
| US-008-05 | Window Default Size | P1 | 5 | 1 | [TASKS-US-008-05.md](FR-008-ui-framework/TASKS-US-008-05.md) |
| US-008-06 | Auto-Start Services | P1 | 6 | 2 | [TASKS-US-008-06.md](FR-008-ui-framework/TASKS-US-008-06.md) |
| US-008-07 | Keyboard Shortcuts | P1 | 5 | 9 | [TASKS-US-008-07.md](FR-008-ui-framework/TASKS-US-008-07.md) |
| US-008-08 | Page Title Header | P1 | 5 | 1 | [TASKS-US-008-08.md](FR-008-ui-framework/TASKS-US-008-08.md) |
| US-008-09 | Clean Shutdown | P1 | 6 | 1 | [TASKS-US-008-09.md](FR-008-ui-framework/TASKS-US-008-09.md) |
| US-008-10 | CJK Font Rendering | P1 | 5 | 1 | [TASKS-US-008-10.md](FR-008-ui-framework/TASKS-US-008-10.md) |

## Legend

- **S** (Small): < 2 hours, single file, straightforward implementation
- **M** (Medium): 2–4 hours, multiple files or complex logic
- **L** (Large): 4–8 hours, architectural decisions, cross-cutting concerns
- **Est** = Effort Estimate
- **Status**: [ ] Not Started, [~] In Progress, [x] Completed
- **—** in Depends On = No dependencies (can start immediately)
