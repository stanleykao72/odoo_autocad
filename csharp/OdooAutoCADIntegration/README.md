# OdooAutoCAD Integration - C# Version

Version 6.0.0 - C# COM Architecture

## Overview

This is the C# implementation of the Odoo-AutoCAD Integration system, migrated from the Python version (v5.1). The application bridges Odoo ERP and AutoCAD for engineering/construction workflows, enabling automated extraction of drawing parameters, BOQ (Bill of Quantities) management, and Purchase Requisition processing.

## Key Features

- **AutoCAD COM Integration**: Full COM interop with AutoCAD (including LT) via late binding
- **GUI Proxy System**: Thread-safe COM operations through message queue architecture
- **MCP SSE Server**: AI assistant integration via Model Context Protocol with 7 tools
- **Odoo REST API**: Complete integration with Odoo ERP for BOQ and PR management
- **WPF MVVM UI**: Modern Windows desktop application with clean architecture

## Architecture

```
OdooAutoCADIntegration/
├── src/
│   ├── OdooAutoCAD.Core/           # Core business logic
│   │   ├── AutoCAD/                # AutoCAD COM service
│   │   ├── Odoo/                   # Odoo REST API client
│   │   └── BOQ/                    # BOQ processing
│   │
│   ├── OdooAutoCAD.Threading/      # Thread-safe COM execution
│   │   └── GUIProxy.cs             # GUI proxy with message queue
│   │
│   ├── OdooAutoCAD.MCP/            # MCP SSE Server
│   │   ├── Server/                 # ASP.NET Core server
│   │   ├── Tools/                  # 7 MCP tools
│   │   └── Protocol/               # JSON-RPC & MCP models
│   │
│   ├── OdooAutoCAD.Data/           # EF Core + SQLite
│   ├── OdooAutoCAD.Configuration/  # YAML/JSON config
│   └── OdooAutoCAD.App/            # WPF application
│
└── tests/                          # xUnit test projects
```

## Technology Stack

| Component | Technology |
|-----------|------------|
| Runtime | .NET 8 (LTS) |
| GUI | WPF + MVVM (CommunityToolkit.Mvvm) |
| AutoCAD | COM Interop (dynamic/late binding) |
| HTTP Server | ASP.NET Core Minimal API |
| HTTP Client | HttpClient |
| Database | SQLite + Entity Framework Core |
| Configuration | YamlDotNet + JSON |
| Logging | Serilog |
| Testing | xUnit + Moq + FluentAssertions |

## Requirements

- Windows 10/11
- .NET 8.0 SDK
- AutoCAD 2020+ (or AutoCAD LT)
- Odoo 14+ with REST API enabled

## Building

```powershell
# Restore packages
dotnet restore

# Build solution
dotnet build

# Run tests
dotnet test

# Run application
dotnet run --project src/OdooAutoCAD.App
```

## Configuration

Edit `appsettings.json`:

```json
{
  "Odoo": {
    "ServerUrl": "https://your-odoo.com",
    "Database": "your_database"
  },
  "MCP": {
    "Port": 8084,
    "AutoStart": false
  }
}
```

## MCP Tools

The application provides 7 MCP tools for AI assistant integration:

1. **test_connection** - Test MCP server connection
2. **get_server_info** - Get server information and capabilities
3. **check_autocad_status** - Check AutoCAD connection status
4. **check_odoo_status** - Check Odoo connection status
5. **extract_autocad_parameters** - Extract parameters from AutoCAD drawings
6. **sync_to_odoo** - Synchronize data to Odoo
7. **generate_boq** - Generate Bill of Quantities

## Threading Model

The application uses a specialized threading architecture to handle AutoCAD COM requirements:

```
GUI Thread (STA)                    MCP Server Thread (MTA)
    │                                       │
    │  ┌─────────────┐                     │
    │  │DispatchTimer│                     │
    │  │   100ms     │                     │
    │  └─────┬───────┘                     │
    │        │                             │
    │        ▼                             │
    │  ┌─────────────┐    Request    ┌─────────────┐
    │  │  GUIProxy   │◄──────────────│ MCP Tools   │
    │  │  Process    │               │             │
    │  └─────┬───────┘               └─────────────┘
    │        │                             ▲
    │        ▼                             │
    │  ┌─────────────┐    Response   ─────┘
    │  │ AutoCAD COM │───────────────
    │  └─────────────┘
```

## Documentation

See `docs/architecture/` for detailed documentation:

- `CSHARP_COM_ARCHITECTURE.md` - Complete architecture design
- `THREADING_MODEL.md` - Threading model details
- `MCP_INTEGRATION.md` - MCP server design
- `MIGRATION_GUIDE.md` - Python to C# migration guide

## License

Copyright (c) 2024 E-Smith. All rights reserved.
