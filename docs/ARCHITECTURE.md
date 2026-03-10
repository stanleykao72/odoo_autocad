# System Architecture

> Version 6.0 — Dual-Mode AutoCAD + MCP Integration

## 1. Overview

Windows desktop application bridging **Odoo ERP** and **AutoCAD** for engineering/construction workflows. Automates extraction of drawing parameters, manages BOQ (Bill of Quantities), and handles Purchase Requisition processing.

```
+------------------+       +-----------------+       +-------------+
|   AutoCAD        | COM   |                 | REST  |   Odoo ERP  |
|   Full 2019+     |<----->|   Python GUI    |<----->|   (Swagger) |
+------------------+       |   (Tkinter)     |       +-------------+
                            |                 |
+------------------+  IPC  |  +-----------+  |
|   AutoCAD LT     |<----->|  | MCP Server|  |
|   2024+          |  File |  | (FastMCP) |  |
+------------------+       +-----------------+
```

## 2. Dual-Mode AutoCAD Connection

| Mode | Backend | AutoCAD Version | Connection |
|------|---------|-----------------|------------|
| **COM** | `util_autocad.py` | Full AutoCAD 2019+ | win32com (pywin32), STA thread |
| **IPC** | `util_autocad_ipc.py` | AutoCAD LT 2024+ | File IPC via autocad-mcp |

**UtilAutoCADDispatcher** (`util_autocad_dispatcher.py`) provides a unified interface.
GUI switches mode at runtime via `COM | IPC` segmented button.

### COM Mode Thread Safety

COM objects must be accessed from the STA (GUI main) thread.
**GUIProxy** (`util_gui_proxy.py`) uses a message queue:

```
MCP Server Thread                    GUI Main Thread
     |                                     |
     |-- execute_in_gui("action") -------->|
     |        (queue request)              |-- process_requests()
     |                                     |     invoke handler
     |<-------- response -----------------|     (COM call here)
     |        (response_cache)             |
```

### IPC Mode Protocol

```
Python                              AutoCAD LT
  |                                      |
  |-- write cmd JSON to %TEMP% --------->|
  |-- PostMessageW (keystrokes) -------->|
  |                                      |-- mcp_dispatch.lsp reads JSON
  |                                      |-- execute command
  |                                      |-- write result JSON
  |<-- poll result JSON -----------------|
```

## 3. Component Map

### Entry Point

| File | Description |
|------|-------------|
| `odoo.py` | CLI entry point. Modes: GUI (default), --mcp-server, --mcp-autocad |

### GUI Layer

| File | Description |
|------|-------------|
| `forms/form_main_modern.py` | Main window (CustomTkinter). Sidebar + content + logs |
| `forms/form_autocad_param.py` | Parameter input dialog |
| `forms/form_autocad_param_enhanced.py` | Enhanced parameter form with Odoo product selection |
| `ui/ui_theme.py` | Color/size/icon constants |
| `ui/ui_fonts.py` | CJK font fallback chain (Microsoft JhengHei UI) |

### Business Logic

| File | Description |
|------|-------------|
| `utility/util_autocad.py` | AutoCAD COM interface (~40 methods) |
| `utility/util_autocad_ipc.py` | AutoCAD LT File IPC client |
| `utility/util_autocad_dispatcher.py` | COM/IPC mode switcher (unified interface) |
| `utility/util_odoo.py` | Odoo Swagger API client |
| `utility/util_push_to_boq.py` | BOQ push workflow orchestrator |
| `utility/util_transfer_boq_to_pr.py` | PR generation workflow orchestrator |
| `utility/util_gui_proxy.py` | GUI thread-safe COM execution (queue-based) |
| `utility/util_log.py` | GUI + console logging |
| `utility/util_load_yaml_config.py` | YAML config loader |

### MCP / AI Assistant

| File | Description |
|------|-------------|
| `mcp_server_autocad.py` | FastMCP server (5 Odoo tools, stdio/SSE) |
| `utility/util_mcp_manager.py` | In-process MCP server lifecycle manager |

### Data Layer

| File | Description |
|------|-------------|
| `models/server.py` | SQLAlchemy ORM: Server table (host, db_name, url, token) |
| `config/*.yaml` | Environment-specific Odoo connection configs |
| `db/database.db` | SQLite (auto-created in %APPDATA%/OdooAutoCAD/) |

### AutoLISP (autolisp/lisp/)

| File | Description |
|------|-------------|
| `main.lsp` | Entry point. Loads modules, registers `OB:MCP-DISPATCH` |
| `ob_mcp_dispatch.lsp` | 6 Odoo IPC actions (extract_tables, write_ids, etc.) |
| `config.lsp` | Path resolution, YAML config locations |
| `table_util.lsp` | 9-column TABLE entity read/write |
| `block_util.lsp` | Attribute block read/write |
| `json_util.lsp` | Pure AutoLISP JSON parser |
| `file_util.lsp` | File I/O utilities |
| `strip_mtext.lsp` | MText formatting cleanup |

### External Dependencies

| Dependency | Usage |
|------------|-------|
| `libs/autocad-mcp/` | Git submodule. File IPC backend + mcp_dispatch.lsp |
| Bravado/Swagger | Odoo REST API client |
| pywin32 | AutoCAD COM (COM mode only) |
| CustomTkinter | Modern GUI widgets |
| FastMCP SDK | MCP protocol server |
| SQLAlchemy | ORM for SQLite config cache |
| uvicorn | SSE transport for MCP server |

## 4. Data Flow

### BOQ Push (Push to Odoo)

```
1. AutoCAD → get_layouts_values()
   - Iterate all layouts (exclude Model)
   - Find 9-column TABLEs with HEADER_ID label
   - Extract rows: position, product_no, width, height, len, thickness, qty, desc

2. Odoo → import2boq(layout_dict)
   - POST /import2boq_v2 with BasicAuth
   - Returns: header_id, detail_id per row

3. AutoCAD → set_layouts_tables_id(boq_list)
   - Write header_id to cell(0, 8)
   - Write detail_id to each data row cell(row, 8)
```

### PR Generation (BOQ → Purchase Requisition)

```
1. AutoCAD → get_layouts_header_id_to_pr()
   - Read header_id from cell(0, 8) of each TABLE

2. Odoo → boq2pr(header_ids)
   - POST /boq2pr_v2
   - Returns: PR reference, lines
```

### TABLE Structure (9 columns)

```
Col:  0          1           2      3       4     5          6    7          8
Row 0: [title................................................] HEADER_ID  {header_id}
Row 1: Position  Product No  Width  Height  Len  Thickness  Qty  Desc       detail_id
Row 2: data...                                                              {detail_id}
Row 3: data...                                                              {detail_id}
```

## 5. Configuration

### Load Priority

1. SQLite DB (%APPDATA%/OdooAutoCAD/database.db)
2. YAML files (c:/odoo/config/server.yaml + token.yaml)
3. Hardcoded defaults

### Server Config Schema

```yaml
# config/server.yaml
host: e-smith.odoo.com
db_name: odoo13-esmith-master-1011507
url: https://e-smith.odoo.com/api/v1/boq_import_api/swagger.json?token=...&db=...

# config/token.yaml
token: 847430a6-...
```

## 6. Odoo API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `import2boq_v2` | PATCH | Push TABLE data → returns header_id/detail_id |
| `boq2pr_v2` | PATCH | Convert BOQ → Purchase Requisition |
| `get_project_v2` | PATCH | Get project info by PR No |
| `get_product_v2` | PATCH | Get product catalog |
| `get_setup_v2` | PATCH | Get setup/configuration values |
| `get_color_v2` | PATCH | Get color options |

All endpoints use BasicAuth: `base64(db_name:token)`.

## 7. MCP Server

5 Odoo-specific tools registered on FastMCP:

| Tool | Description |
|------|-------------|
| `odoo_push_boq` | Extract TABLEs + push to Odoo + writeback IDs |
| `odoo_create_pr` | Read header_ids + create Purchase Requisition |
| `odoo_get_setup` | Fetch product/spec/config from Odoo |
| `odoo_get_colors` | Fetch color options |
| `odoo_status` | Check AutoCAD + Odoo connection status |

Transport: stdio (CLI) or SSE (port 8084, in-process via MCPManager).

## 8. Deployment

### VLX Packaging (AutoCAD side)

Two VLX files built via `autolisp/build_vlx.lsp`:

| File | Contents |
|------|----------|
| `McpDispatch.vlx` | mcp_dispatch.lsp + attribute_tools.lsp (IPC core) |
| `OdooAutoCAD.vlx` | 8 Odoo .lsp modules (TABLE/block/config/JSON) |

Load order: McpDispatch.vlx first, then OdooAutoCAD.vlx.

### Python Application

```
python odoo.py                          # GUI (COM mode, default)
python odoo.py --autocad-mode ipc       # GUI (IPC mode)
python odoo.py --mcp-autocad            # MCP stdio server
python odoo.py --enable-mcp             # GUI + MCP auto-start
```
