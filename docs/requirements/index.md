# C# v6.0 Feature Requirements Index

> **Version**: 2.0
> **Created**: 2026-02-06
> **Purpose**: Master index for all feature requirement documents and user stories for the C# WPF migration

## Overview

This directory contains comprehensive feature requirement documents and granular user stories for every page and feature in the C# v6.0 WPF application. Each FR document maps Python v5.1 functionality to C# implementation targets, and each US file provides actionable acceptance criteria with implementation tasks.

**Total**: 11 Feature Requirements | 87 User Stories | 279 Functional Requirements

## Document Structure

```
docs/requirements/
├── index.md                              ← You are here
├── FR-001-dashboard/                     (4 user stories)
├── FR-002-autocad-connection/            (7 user stories)
├── FR-003-odoo-connection/               (12 user stories)
├── FR-004-boq-manager/                   (12 user stories)
├── FR-005-purchase-requisition/          (8 user stories)
├── FR-006-ai-assistant/                  (12 user stories)
├── FR-007-settings/                      (13 user stories)
├── FR-008-ui-framework/                  (10 user stories)
├── FR-009-parameter-config/              (3 user stories)
├── FR-010-dual-mode-autocad/             (4 user stories)
└── FR-011-table-entity-writer/           (2 user stories)
```

## Feature Requirements

| ID | Document | Feature | US Count | Status | Priority |
|----|----------|---------|----------|--------|----------|
| FR-001 | [FR-001-dashboard.md](FR-001-dashboard/FR-001-dashboard.md) | Dashboard Page | 4 | Not Started | P1 |
| FR-002 | [FR-002-autocad-connection.md](FR-002-autocad-connection/FR-002-autocad-connection.md) | AutoCAD Integration Page | 7 | Not Started | P1 |
| FR-003 | [FR-003-odoo-connection.md](FR-003-odoo-connection/FR-003-odoo-connection.md) | Odoo Integration Page | 12 | Not Started | P1 |
| FR-004 | [FR-004-boq-manager.md](FR-004-boq-manager/FR-004-boq-manager.md) | BOQ Manager Page | 12 | Not Started | P2 |
| FR-005 | [FR-005-purchase-requisition.md](FR-005-purchase-requisition/FR-005-purchase-requisition.md) | Purchase Requisition Page | 8 | Not Started | P2 |
| FR-006 | [FR-006-ai-assistant.md](FR-006-ai-assistant/FR-006-ai-assistant.md) | AI Assistant (MCP) Page | 12 | Not Started | P3 |
| FR-007 | [FR-007-settings.md](FR-007-settings/FR-007-settings.md) | Settings Page | 13 | Not Started | P2 |
| FR-008 | [FR-008-ui-framework.md](FR-008-ui-framework/FR-008-ui-framework.md) | UI Framework & Navigation | 10 | Not Started | P1 |
| FR-009 | [FR-009-parameter-config.md](FR-009-parameter-config/FR-009-parameter-config.md) | Odoo Parameter Configuration | 3 | Not Started | P2 |
| FR-010 | [FR-010-dual-mode-autocad.md](FR-010-dual-mode-autocad/FR-010-dual-mode-autocad.md) | Dual-Mode AutoCAD Support | 4 | Not Started | P1 |
| FR-011 | [FR-011-table-entity-writer.md](FR-011-table-entity-writer/FR-011-table-entity-writer.md) | TABLE Entity Writer | 2 | Not Started | P1 |

---

## User Stories by Feature

### FR-001: Dashboard Page (4 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-001-01](FR-001-dashboard/US-001-01-view-connection-status.md) | View Connection Status | P1 | US-008-01 |
| [US-001-02](FR-001-dashboard/US-001-02-quick-connect-buttons.md) | Quick Connect Buttons | P1 | US-001-01 |
| [US-001-03](FR-001-dashboard/US-001-03-recent-activity.md) | Recent Activity | P2 | US-001-01 |
| [US-001-04](FR-001-dashboard/US-001-04-mcp-server-status.md) | MCP Server Status | P3 | US-001-01, US-006-01 |

### FR-002: AutoCAD Integration Page (7 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-002-01](FR-002-autocad-connection/US-002-01-connect-to-autocad.md) | Connect to AutoCAD | P1 | US-008-01 |
| [US-002-02](FR-002-autocad-connection/US-002-02-view-layouts.md) | View Layouts | P1 | US-002-01 |
| [US-002-03](FR-002-autocad-connection/US-002-03-extract-parameters.md) | Extract Parameters | P1 | US-002-01 |
| [US-002-04](FR-002-autocad-connection/US-002-04-drawing-info.md) | Drawing Info | P2 | US-002-01 |
| [US-002-05](FR-002-autocad-connection/US-002-05-pr-project-info.md) | PR Project Info | P2 | US-002-01, US-003-01 |
| [US-002-06](FR-002-autocad-connection/US-002-06-clear-table-ids.md) | Clear Table IDs | P2 | US-002-03 |
| [US-002-07](FR-002-autocad-connection/US-002-07-monitor-com-status.md) | Monitor COM Status | P2 | US-002-01 |

### FR-003: Odoo Integration Page (12 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-003-01](FR-003-odoo-connection/US-003-01-enter-odoo-credentials.md) | Enter Odoo Credentials | P1 | US-008-01 |
| [US-003-02](FR-003-odoo-connection/US-003-02-test-connection.md) | Test Connection | P1 | US-003-01 |
| [US-003-03](FR-003-odoo-connection/US-003-03-view-connection-status.md) | View Connection Status | P1 | US-003-01 |
| [US-003-04](FR-003-odoo-connection/US-003-04-disconnect-from-odoo.md) | Disconnect from Odoo | P1 | US-003-01 |
| [US-003-05](FR-003-odoo-connection/US-003-05-sync-product-catalog.md) | Sync Product Catalog | P1 | US-003-02 |
| [US-003-06](FR-003-odoo-connection/US-003-06-search-products.md) | Search Products | P2 | US-003-05 |
| [US-003-07](FR-003-odoo-connection/US-003-07-search-projects.md) | Search Projects | P2 | US-003-02 |
| [US-003-08](FR-003-odoo-connection/US-003-08-view-last-sync-time.md) | View Last Sync Time | P2 | US-003-05 |
| [US-003-09](FR-003-odoo-connection/US-003-09-view-server-info.md) | View Server Info | P2 | US-003-02 |
| [US-003-10](FR-003-odoo-connection/US-003-10-configure-via-appsettings.md) | Configure via appsettings.json | P1 | — |
| [US-003-11](FR-003-odoo-connection/US-003-11-filter-by-category.md) | Filter by Category | P2 | US-003-05 |
| [US-003-12](FR-003-odoo-connection/US-003-12-persist-credentials.md) | Persist Credentials | P1 | US-003-01 |

### FR-004: BOQ Manager Page (12 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-004-01](FR-004-boq-manager/US-004-01-extract-boq-data.md) | Extract BOQ Data | P1 | US-002-01 |
| [US-004-02](FR-004-boq-manager/US-004-02-review-extracted-data.md) | Review Extracted Data | P1 | US-004-01 |
| [US-004-03](FR-004-boq-manager/US-004-03-validate-against-products.md) | Validate Against Products | P1 | US-004-01, US-003-05 |
| [US-004-04](FR-004-boq-manager/US-004-04-push-to-odoo.md) | Push to Odoo | P1 | US-004-03 |
| [US-004-05](FR-004-boq-manager/US-004-05-id-writeback.md) | ID Writeback | P1 | US-004-04 |
| [US-004-06](FR-004-boq-manager/US-004-06-view-skipped-rows.md) | View Skipped Rows | P2 | US-004-01 |
| [US-004-07](FR-004-boq-manager/US-004-07-view-summary.md) | View Summary | P2 | US-004-04 |
| [US-004-08](FR-004-boq-manager/US-004-08-map-products.md) | Map Products | P2 | US-004-03 |
| [US-004-09](FR-004-boq-manager/US-004-09-view-validation-errors.md) | View Validation Errors | P2 | US-004-03 |
| [US-004-10](FR-004-boq-manager/US-004-10-clear-ids.md) | Clear IDs | P2 | US-002-01 |
| [US-004-11](FR-004-boq-manager/US-004-11-manage-mappings.md) | Manage Mappings | P2 | US-004-08 |
| [US-004-12](FR-004-boq-manager/US-004-12-progress-indication.md) | Progress Indication | P2 | US-004-04 |

### FR-005: Purchase Requisition Page (8 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-005-01](FR-005-purchase-requisition/US-005-01-convert-boq-to-pr.md) | Convert BOQ to PR | P1 | US-004-04 |
| [US-005-02](FR-005-purchase-requisition/US-005-02-view-pr-list.md) | View PR List | P1 | US-005-01 |
| [US-005-03](FR-005-purchase-requisition/US-005-03-view-pr-details.md) | View PR Details | P2 | US-005-02 |
| [US-005-04](FR-005-purchase-requisition/US-005-04-submit-pr.md) | Submit PR | P1 | US-005-01 |
| [US-005-05](FR-005-purchase-requisition/US-005-05-track-pr-status.md) | Track PR Status | P2 | US-005-02 |
| [US-005-06](FR-005-purchase-requisition/US-005-06-filter-and-sort-prs.md) | Filter and Sort PRs | P2 | US-005-02 |
| [US-005-07](FR-005-purchase-requisition/US-005-07-conversion-feedback.md) | Conversion Feedback | P2 | US-005-01 |
| [US-005-08](FR-005-purchase-requisition/US-005-08-view-pr-totals.md) | View PR Totals | P2 | US-005-02 |

### FR-006: AI Assistant / MCP Page (12 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-006-01](FR-006-ai-assistant/US-006-01-start-mcp-server.md) | Start MCP Server | P3 | US-008-01 |
| [US-006-02](FR-006-ai-assistant/US-006-02-stop-mcp-server.md) | Stop MCP Server | P3 | US-006-01 |
| [US-006-03](FR-006-ai-assistant/US-006-03-view-server-status.md) | View Server Status | P3 | US-006-01 |
| [US-006-04](FR-006-ai-assistant/US-006-04-test-connection.md) | Test Connection | P3 | US-006-01 |
| [US-006-05](FR-006-ai-assistant/US-006-05-view-tool-registry.md) | View Tool Registry | P3 | US-006-01 |
| [US-006-06](FR-006-ai-assistant/US-006-06-view-configuration.md) | View Configuration | P3 | US-006-01 |
| [US-006-07](FR-006-ai-assistant/US-006-07-monitor-connections.md) | Monitor Connections | P3 | US-006-01 |
| [US-006-08](FR-006-ai-assistant/US-006-08-view-activity-logs.md) | View Activity Logs | P3 | US-006-01 |
| [US-006-09](FR-006-ai-assistant/US-006-09-view-tool-prerequisites.md) | View Tool Prerequisites | P3 | US-006-05 |
| [US-006-10](FR-006-ai-assistant/US-006-10-restart-server.md) | Restart Server | P3 | US-006-01, US-006-02 |
| [US-006-11](FR-006-ai-assistant/US-006-11-view-tool-execution.md) | View Tool Execution | P3 | US-006-05 |
| [US-006-12](FR-006-ai-assistant/US-006-12-configure-server-port.md) | Configure Server Port | P3 | US-006-01 |

### FR-007: Settings Page (13 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-007-01](FR-007-settings/US-007-01-configure-odoo-connection.md) | Configure Odoo Connection | P1 | US-008-01 |
| [US-007-02](FR-007-settings/US-007-02-switch-environments.md) | Switch Environments | P2 | US-007-01 |
| [US-007-03](FR-007-settings/US-007-03-test-odoo-connection.md) | Test Odoo Connection | P2 | US-007-01 |
| [US-007-04](FR-007-settings/US-007-04-autocad-timeout-settings.md) | AutoCAD Timeout Settings | P2 | US-007-01 |
| [US-007-05](FR-007-settings/US-007-05-mcp-port-configuration.md) | MCP Port Configuration | P3 | US-007-01 |
| [US-007-06](FR-007-settings/US-007-06-mcp-auto-start-toggle.md) | MCP Auto-Start Toggle | P3 | US-007-01 |
| [US-007-07](FR-007-settings/US-007-07-change-theme.md) | Change Theme | P2 | US-008-01 |
| [US-007-08](FR-007-settings/US-007-08-set-language.md) | Set Language | P3 | US-007-01 |
| [US-007-09](FR-007-settings/US-007-09-change-log-level.md) | Change Log Level | P2 | US-007-01 |
| [US-007-10](FR-007-settings/US-007-10-clear-cache.md) | Clear Cache | P2 | US-007-01 |
| [US-007-11](FR-007-settings/US-007-11-export-configuration.md) | Export Configuration | P2 | US-007-01 |
| [US-007-12](FR-007-settings/US-007-12-import-configuration.md) | Import Configuration | P2 | US-007-11 |
| [US-007-13](FR-007-settings/US-007-13-view-app-info.md) | View App Info | P3 | US-008-01 |

### FR-008: UI Framework & Navigation (10 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-008-01](FR-008-ui-framework/US-008-01-sidebar-navigation.md) | Sidebar Navigation | P1 | — |
| [US-008-02](FR-008-ui-framework/US-008-02-active-page-indicator.md) | Active Page Indicator | P1 | US-008-01 |
| [US-008-03](FR-008-ui-framework/US-008-03-persistent-connection-status.md) | Persistent Connection Status | P1 | US-008-01 |
| [US-008-04](FR-008-ui-framework/US-008-04-system-log-panel.md) | System Log Panel | P2 | US-008-01 |
| [US-008-05](FR-008-ui-framework/US-008-05-window-default-size.md) | Window Default Size | P1 | US-008-01 |
| [US-008-06](FR-008-ui-framework/US-008-06-auto-start-services.md) | Auto-Start Services | P2 | US-008-01 |
| [US-008-07](FR-008-ui-framework/US-008-07-keyboard-shortcuts.md) | Keyboard Shortcuts | P3 | US-008-01 |
| [US-008-08](FR-008-ui-framework/US-008-08-page-title-header.md) | Page Title Header | P1 | US-008-01 |
| [US-008-09](FR-008-ui-framework/US-008-09-clean-shutdown.md) | Clean Shutdown | P1 | US-008-01 |
| [US-008-10](FR-008-ui-framework/US-008-10-cjk-font-rendering.md) | CJK Font Rendering | P1 | US-008-01 |

### FR-009: Odoo Parameter Configuration (3 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-009-01](FR-009-parameter-config/US-009-01-fetch-odoo-parameters.md) | Fetch Odoo Parameters | P2 | US-003-02, US-003-05 |
| [US-009-02](FR-009-parameter-config/US-009-02-parameter-form-ui.md) | Parameter Form UI | P2 | US-009-01 |
| [US-009-03](FR-009-parameter-config/US-009-03-write-parameters-to-autocad.md) | Write Parameters to AutoCAD | P2 | US-009-02, US-002-01 |

### FR-010: Dual-Mode AutoCAD Support (4 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-010-01](FR-010-dual-mode-autocad/US-010-01-select-operation-mode.md) | Select Operation Mode | P1 | US-007-01, US-002-01 |
| [US-010-02](FR-010-dual-mode-autocad/US-010-02-file-based-extraction.md) | File-Based Data Extraction | P1 | US-010-01 |
| [US-010-03](FR-010-dual-mode-autocad/US-010-03-file-based-writeback.md) | File-Based Writeback | P1 | US-010-02 |
| [US-010-04](FR-010-dual-mode-autocad/US-010-04-mode-switch-runtime.md) | Runtime Mode Switching | P1 | US-010-01, US-010-02, US-010-03 |

### FR-011: TABLE Entity Writer (2 stories)

| ID | Title | Priority | Depends On |
|----|-------|----------|------------|
| [US-011-01](FR-011-table-entity-writer/US-011-01-dxf-table-writer.md) | DXF TABLE Entity Writer | P1 | US-010-03 |
| [US-011-02](FR-011-table-entity-writer/US-011-02-dwg-table-writer.md) | DWG TABLE Entity Writer | P2 | US-011-01 |

---

## Implementation Priority

### Phase 1 - Core Infrastructure (P1)
1. **FR-008** UI Framework & Navigation - foundation for all pages
2. **FR-001** Dashboard - entry point, connection status overview
3. **FR-002** AutoCAD Integration - core COM connection
4. **FR-003** Odoo Integration - core API connection

### Phase 2 - Business Logic (P2)
5. **FR-004** BOQ Manager - primary business workflow
6. **FR-005** Purchase Requisition - secondary business workflow
7. **FR-007** Settings - configuration management

### Phase 3.5 - Dual-Mode & TABLE Writing (P1)
8. **FR-010** Dual-Mode AutoCAD - COM + File mode support
9. **FR-011** TABLE Entity Writer - native ID writeback in File mode

### Phase 4 - Advanced Features (P3)
10. **FR-006** AI Assistant (MCP) - AI integration features

### Recommended Sprint Order (P1 stories)

| Sprint | Stories | Description |
|--------|---------|-------------|
| 1 | US-008-01, US-008-02, US-008-05, US-008-08, US-008-09, US-008-10 | Shell app: navigation, window, shutdown, fonts |
| 2 | US-008-03, US-001-01, US-001-02 | Dashboard with connection status |
| 3 | US-002-01, US-002-02, US-002-03 | AutoCAD connect, layouts, extract |
| 4 | US-003-01, US-003-02, US-003-03, US-003-04, US-003-10, US-003-12 | Odoo credentials, connect, config |
| 5 | US-003-05, US-004-01, US-004-02, US-004-03 | Product sync, BOQ extract & validate |
| 6 | US-004-04, US-004-05, US-005-01, US-005-04 | Push BOQ, writeback, PR conversion |

---

## C# Project Structure Reference

```
csharp/OdooAutoCADIntegration/src/
├── OdooAutoCAD.App/           # WPF Application
│   ├── Views/
│   │   ├── MainWindow.xaml     → FR-008
│   │   └── Pages/
│   │       ├── DashboardPage.xaml        → FR-001
│   │       ├── AutoCADPage.xaml          → FR-002
│   │       ├── OdooPage.xaml             → FR-003
│   │       ├── BOQPage.xaml              → FR-004
│   │       ├── PurchaseRequisitionPage.xaml → FR-005
│   │       ├── MCPPage.xaml              → FR-006
│   │       └── SettingsPage.xaml         → FR-007
│   ├── ViewModels/
│   │   ├── MainViewModel.cs              → FR-008
│   │   ├── DashboardViewModel.cs         → FR-001
│   │   ├── AutoCADViewModel.cs           → FR-002
│   │   ├── OdooViewModel.cs              → FR-003
│   │   ├── BOQViewModel.cs               → FR-004
│   │   ├── PurchaseRequisitionViewModel.cs → FR-005
│   │   ├── MCPViewModel.cs               → FR-006
│   │   └── SettingsViewModel.cs          → FR-007
│   └── Services/NavigationService.cs     → FR-008
├── OdooAutoCAD.Core/          # Business Logic
│   ├── AutoCAD/IAutoCADService.cs        → FR-002
│   ├── AutoCAD/IDrawingDataService.cs   → FR-010
│   ├── AutoCAD/IDwgFileService.cs       → FR-010
│   ├── Odoo/IOdooService.cs              → FR-003
│   ├── BOQ/IBOQProcessor.cs              → FR-004
│   └── Threading/IGUIProxy.cs            → FR-006, FR-008
├── OdooAutoCAD.MCP/           # MCP Server
│   ├── Server/MCPSSEServer.cs            → FR-006
│   ├── Tools/MCPToolRegistry.cs          → FR-006
│   └── Protocol/MCPModels.cs             → FR-006
├── OdooAutoCAD.Data/          # Data Layer
│   ├── Context/AppDbContext.cs           → FR-007
│   └── Entities/Entities.cs             → FR-003, FR-004
├── OdooAutoCAD.Threading/     # Thread Safety
│   └── GUIProxy.cs                      → FR-006, FR-008
└── OdooAutoCAD.Configuration/ # Configuration
    └── ConfigurationLoader.cs           → FR-007
```

## Python Source Reference

| Python File | Lines | Mapped To |
|-------------|-------|-----------|
| `forms/form_main_modern.py` | ~570 | FR-001 through FR-008 |
| `utility/util_autocad.py` | ~1,301 | FR-002 |
| `utility/util_odoo.py` | ~500+ | FR-003 |
| `utility/util_push_to_boq.py` | ~100 | FR-004 |
| `utility/util_transfer_boq_to_pr.py` | ~50 | FR-005 |
| `mcp_server_fastmcp.py` | ~1,500 | FR-006 |
| `utility/util_mcp_sse_manager.py` | ~925 | FR-006 |
| `utility/util_gui_proxy.py` | ~200+ | FR-006, FR-008 |
| `config/*.yaml` | N/A | FR-007 |
| `ui/ui_theme.py`, `ui/ui_fonts.py` | N/A | FR-008 |

## Document Template

Each FR document follows this consistent structure:

1. **Overview** - Feature purpose and scope
2. **User Stories** - Who uses it, what they do, why
3. **Python Reference** - Source files and key functions
4. **Functional Requirements** - Numbered requirements (FR-XXX-NNN)
5. **UI Wireframe Description** - Layout, controls, interactions
6. **Data Model** - Input/output data structures, C# models
7. **API/Service Dependencies** - C# interfaces and services needed
8. **Validation Rules** - Input validation, business rules
9. **Error Handling** - Error scenarios and user-facing messages
10. **Implementation Notes** - C# target files, MVVM bindings, special considerations

Each US document follows this structure:

1. **User Story** - As a / I want to / So that
2. **Parent Feature** - FR link and priority
3. **Acceptance Criteria** - Testable checkbox items
4. **Related Functional Requirements** - FR-XXX-NNN mapping table
5. **Implementation Tasks** - Task ID, description, target file, estimate
6. **Dependencies** - Depends on / Blocks
7. **Notes** - Additional implementation context
