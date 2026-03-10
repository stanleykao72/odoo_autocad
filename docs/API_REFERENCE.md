# API Reference

## 1. Python Public Interfaces

### UtilAutoCADDispatcher

Unified interface for COM and IPC backends. All downstream code should use this.

```python
from utility.util_autocad_dispatcher import UtilAutoCADDispatcher

dispatcher = UtilAutoCADDispatcher(odoo_util, log_util, mode="com")
```

| Method | Returns | Description |
|--------|---------|-------------|
| `connected_autocad()` | `bool` | Connection status |
| `connect_autocad(main_body)` | — | Establish connection |
| `switch_mode(new_mode)` | — | Switch "com" ↔ "ipc" at runtime |
| `mode` | `str` | Current mode ("com" or "ipc") |
| `get_layouts_values()` | `dict` | Extract all TABLE data from all layouts |
| `set_layouts_tables_id(boq_list)` | — | Write header_id/detail_id back |
| `get_layouts_header_id_to_pr()` | `list` | Collect header_ids for PR |
| `get_block_attributes()` | `dict` | Read attribute block values |
| `set_block_attributes(attrs, layout_name)` | — | Write attribute block values |
| `get_active_layout()` | `object` | Current layout |
| `get_doc_layouts()` | `list` | All layouts (exclude Model) |
| `clear_table_id(layout)` | — | Clear IDs in one layout |
| `clear_all_tables_id()` | — | Clear IDs in all layouts |
| `draw_line(start, end, layer)` | — | Draw a line |
| `draw_circle(center, radius, layer)` | — | Draw a circle |
| `set_layer(name, color, create)` | — | Set/create layer |
| `list_layers(filter, sort, details)` | `list` | List layers |

### UtilOdoo

Odoo REST API client via Bravado/Swagger.

```python
from utility.util_odoo import UtilOdoo

odoo_util = UtilOdoo(odoo_connection, log_util)
```

| Method | Returns | Description |
|--------|---------|-------------|
| `connected_odoo()` | `bool` | Check Swagger client exists |
| `connect_odoo(conn)` | `(odoo, opts, token)` | Connect via Swagger |
| `import2boq(layout_dict)` | `list` | Push TABLE data, get IDs back |
| `boq2pr(layout_dict)` | `list` | Convert BOQ → Purchase Requisition |
| `get_project(pr_no)` | `dict` | Get project by PR No |
| `get_product(domain)` | `list` | Get product catalog |
| `get_setup(project_id)` | `dict` | Get setup/config values |
| `get_color(domain)` | `list` | Get color options |

### GUIProxy

Thread-safe COM execution via message queue. COM mode only.

```python
from utility.util_gui_proxy import setup_gui_proxy_handlers

proxy = setup_gui_proxy_handlers(autocad_util, log_util)
```

| Method | Description |
|--------|-------------|
| `execute_in_gui(action, **kwargs)` | Queue request, block until response (10s timeout) |
| `process_requests()` | Called by GUI mainloop timer (100ms). Invokes handlers on STA thread |
| `register_handler(name, func)` | Register a handler function |

**Built-in handlers**: switch_layout, get_current_layout, extract_parameters, extract_layout_parameters, get_autocad_status, draw_line, draw_circle, export_layout_image

### MCPManager

In-process MCP server lifecycle.

```python
from utility.util_mcp_manager import MCPManager

mgr = MCPManager(autocad_dispatcher, odoo_util, transport="sse", port=8084)
mgr.start_server()
```

| Method | Description |
|--------|-------------|
| `start_server()` | Start uvicorn SSE server in background thread |
| `stop_server()` | Graceful shutdown |
| `restart_server()` | Stop + wait 1s + start |
| `set_status_callback(fn)` | `fn(is_running: bool, message: str)` |
| `update_autocad_dispatcher(d)` | Update after mode switch |

## 2. Odoo REST API

Base: `https://{host}/api/v1/boq_import_api/swagger.json`

Auth: BasicAuth `base64(db_name:token)`

### import2boq_v2

Push TABLE data to Odoo, receive header_id/detail_id.

```
PATCH /api/v1/boq_import_api/import2boq_v2

Body: {
  "layout_dict": {
    "Layout1": {
      "header": { "pr_no": "...", "project_name": "..." },
      "details": [
        { "position": "1", "product_no": "A001", "qty": 10, ... }
      ]
    }
  }
}

Response: [
  { "header_id": 123, "details": [{ "detail_id": 456, ... }] }
]
```

### boq2pr_v2

Convert BOQ headers to Purchase Requisition.

```
PATCH /api/v1/boq_import_api/boq2pr_v2

Body: { "header_ids": [123, 456] }

Response: [
  { "pr_id": 789, "pr_no": "PR-001", "lines": [...] }
]
```

### get_project_v2 / get_product_v2 / get_setup_v2 / get_color_v2

```
PATCH /api/v1/boq_import_api/{endpoint}

Body: { "args": [domain_filter], "user_token": "..." }

Response: { ... }
```

## 3. MCP Tools (mcp_server_autocad.py)

| Tool | Parameters | Description |
|------|-----------|-------------|
| `odoo_push_boq` | `project_id: int = 0` | Extract TABLEs → push Odoo → writeback IDs |
| `odoo_create_pr` | — | Read header_ids → create PR |
| `odoo_get_setup` | `project_id: int = 0` | Fetch setup/config values |
| `odoo_get_colors` | — | Fetch color options |
| `odoo_status` | — | Check AutoCAD + Odoo connection |

## 4. AutoLISP IPC Actions (070_ob_mcp_dispatch.lsp)

Dispatched via `mcp-dispatch-command`:

| Command | Parameters | Description |
|---------|-----------|-------------|
| `odoo_extract_tables` | — | Get all TABLE data as JSON |
| `odoo_get_header_ids` | — | Get header_ids from all TABLEs |
| `odoo_write_ids` | `{boq_list: [...]}` | Write header_id/detail_id to TABLEs |
| `odoo_get_block_attrs` | — | Read attribute block values |
| `odoo_set_block_attrs` | `{attrs: {...}, layout: "..."}` | Write attribute block values |
| `odoo_clear_ids` | — | Clear all IDs from TABLEs |

JSON format:
```json
// Command (Python → AutoCAD)
{"request_id": "uuid", "command": "odoo_extract_tables", "params": {}}

// Result (AutoCAD → Python)
{"request_id": "uuid", "ok": true, "payload": {...}}
```

## 5. Database Schema

### Server Table (SQLite)

```sql
CREATE TABLE server (
    id       INTEGER PRIMARY KEY,
    host     VARCHAR,      -- e.g. "e-smith.odoo.com"
    db_name  VARCHAR,      -- e.g. "odoo13-esmith-master-1011507"
    url      VARCHAR,      -- Swagger JSON endpoint
    token    VARCHAR,      -- BasicAuth token
    sync_yaml BOOLEAN DEFAULT 0,
    active   BOOLEAN DEFAULT 1
);
```

Location: `%APPDATA%/OdooAutoCAD/database.db`
