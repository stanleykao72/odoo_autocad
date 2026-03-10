# Workflows

## 1. Application Startup

```
python odoo.py [--autocad-mode com|ipc] [--enable-mcp]
  │
  ├─ sqlite_create_table()
  │    Load Server config from SQLite
  │    Fallback: YAML (c:/odoo/config/) → defaults
  │
  ├─ ModernFormMain(odoo_connection, autocad_mode)
  │    ├─ create_ui()          → top bar, sidebar, main area, log
  │    ├─ setup_log_util()     → ScrolledText or console
  │    ├─ init_utilities()
  │    │    ├─ UtilOdoo         → Swagger client
  │    │    ├─ UtilAutoCADDispatcher(mode)
  │    │    │    ├─ COM: UtilAutoCAD (lazy)
  │    │    │    └─ IPC: UtilAutoCADIPC (lazy)
  │    │    ├─ MCPManager       → shared instances configured
  │    │    ├─ GUIProxy         → COM mode only (queue + 100ms timer)
  │    │    └─ auto_start_mcp_server()  → uvicorn SSE on port 8084
  │    └─ mainloop()
  │
  └─ Exit: stop MCP server → destroy window
```

## 2. Connect to AutoCAD (COM Mode)

```
User clicks "連接到 AutoCAD"
  │
  ├─ UtilAutoCAD.connect_autocad()
  │    ├─ pythoncom.CoInitialize()
  │    ├─ GetActiveObject("AutoCAD.Application")
  │    │    or Dispatch (launch new instance)
  │    ├─ Retry ActiveDocument 5x (1s interval)
  │    └─ process_pr_no() → read block attrs → get_project from Odoo
  │
  └─ update_connection_status()
       └─ Top bar: "AutoCAD (COM): ✅ 已連接"
```

## 3. Connect to AutoCAD (IPC Mode)

```
User clicks "連接到 AutoCAD LT"
  │
  ├─ UtilAutoCADIPC.connect_autocad()
  │    ├─ FileIPCBackend() from autocad-mcp
  │    ├─ backend.status() → ping via JSON file exchange
  │    └─ Set connected = True
  │
  └─ Requires: AutoCAD LT running with mcp_dispatch.lsp loaded
```

## 4. Push to BOQ

```
User clicks "推送到 BOQ"
  │
  ├─ UtilPushToBoq.push_to_boq()
  │    │
  │    ├─ Step 1: autocad_util.get_layouts_values()
  │    │    ├─ Iterate all layouts (exclude Model)
  │    │    ├─ Find 9-column TABLEs (chk_legal_table)
  │    │    └─ Extract rows → layout_dict
  │    │
  │    ├─ Step 2: odoo_util.import2boq(layout_dict)
  │    │    ├─ POST /import2boq_v2 (BasicAuth)
  │    │    └─ Returns boq_list with header_id, detail_id
  │    │
  │    └─ Step 3: autocad_util.set_layouts_tables_id(boq_list)
  │         ├─ Write header_id → cell(0, 8)
  │         └─ Write detail_id → cell(row, 8) for each data row
  │
  └─ Log: "BOQ pushed successfully"
```

## 5. Transfer BOQ to PR

```
User clicks "轉移到 PR"
  │
  ├─ UtilTransferBoqToPr.transfer_boq_to_pr()
  │    │
  │    ├─ Step 1: autocad_util.get_layouts_header_id_to_pr()
  │    │    └─ Read header_id from cell(0, 8) of each TABLE
  │    │
  │    └─ Step 2: odoo_util.boq2pr(header_ids)
  │         └─ POST /boq2pr_v2 → returns PR reference
  │
  └─ Log: "PR created"
```

## 6. Get Parameters from Odoo

```
User clicks "取得參數"
  │
  ├─ EnhancedFormAutoCADParam.get_parameters_from_odoo()
  │    ├─ odoo_util.get_setup() → products, specs, operations
  │    ├─ odoo_util.get_color() → color options
  │    ├─ Show selection form (dropdowns)
  │    └─ autocad_util.set_block_attributes(selected_values)
  │         └─ Write to attribute block in all layouts
  │
  └─ Tags written: product_name, spec, product_catelog,
     operation_flow, surface_treatment, color_name, color_no
```

## 7. Mode Switch (Runtime)

```
User clicks COM|IPC segmented button
  │
  ├─ _on_mode_switch("ipc")
  │    ├─ autocad_dispatcher.switch_mode("ipc")
  │    │    └─ _active = UtilAutoCADIPC (lazy init)
  │    ├─ gui_proxy = None (IPC doesn't need COM proxy)
  │    ├─ mcp_manager.update_autocad_dispatcher()
  │    └─ Update button text: "連接到 AutoCAD LT"
  │
  └─ Status bar: "AutoCAD (IPC): ❌ 未連接"
```

## 8. MCP Tool Execution (AI Assistant)

```
AI Client (Claude/Gemini)
  │
  ├─ MCP request → odoo_push_boq(project_id=0)
  │    │
  │    ├─ mcp_server_autocad.py receives call
  │    ├─ Uses global _autocad_dispatcher
  │    │    └─ Same instance as GUI (shared via MCPManager)
  │    ├─ COM mode: GUIProxy queues COM calls
  │    │    IPC mode: direct File IPC (no proxy needed)
  │    └─ Returns JSON result
  │
  └─ AI Client displays result
```

## 9. VLX Build

```
Full AutoCAD command line:
  (load "C:/odoo/autocad_source/autolisp/build_vlx.lsp")
  BUILD-ALL
  │
  ├─ BUILD-MCP-DISPATCH
  │    ├─ Compile attribute_tools.lsp → .fas
  │    ├─ Compile mcp_dispatch.lsp → .fas
  │    └─ VLISP-MAKE-APP → dist/McpDispatch.vlx
  │
  └─ BUILD-ODOO-VLX
       ├─ Compile 8 .lsp files → .fas (dependency order)
       └─ VLISP-MAKE-APP → dist/OdooAutoCAD.vlx
```
