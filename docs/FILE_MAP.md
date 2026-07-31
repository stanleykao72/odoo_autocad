# File Map

Complete listing of all source files and their roles.

## Root

```
odoo.py                        Entry point (CLI + GUI launcher)
mcp_server_autocad.py          MCP server (13 tools: 8 autocad-mcp + 5 Odoo)
CLAUDE.md                      Project instructions for Claude Code
version.py                     Single source of truth for APP_VERSION
pytest.ini                     Test config (live AutoCAD tests deselected by default)
```

## forms/ — GUI

```
form_main_modern.py            Main window (CustomTkinter), sidebar navigation
form_autocad_param.py          Parameter input dialog (basic)
form_autocad_param_enhanced.py Parameter input with Odoo product selection
```

## utility/ — Business Logic

```
util_autocad.py                AutoCAD COM interface (Full AutoCAD)
util_autocad_ipc.py            AutoCAD File IPC client (LT 2024+)
util_autocad_dispatcher.py     COM/IPC mode switcher (unified API)
util_odoo.py                   Odoo REST API client (Swagger/Bravado)
util_push_to_boq.py            BOQ push workflow orchestrator
util_transfer_boq_to_pr.py     PR generation workflow orchestrator
util_gui_proxy.py              Thread-safe COM execution (queue-based)
util_mcp_manager.py            MCP server lifecycle manager (in-process)
util_log.py                    GUI + console logger
util_load_yaml_config.py       YAML configuration loader
util_secrets.py                Secret masking for log output
popup_selector.py              Popup selection dialog
```

## utility/ — ABC contract

```
autocad_backend_interface.py   ABC: 17 abstract methods both backends implement
```

## models/ — Data

```
server.py                      SQLAlchemy ORM: Server table
```

## ui/ — Theme & Fonts

```
ui_theme.py                    Colors, sizes, icons
ui_fonts.py                    CJK font fallback chain
enhanced_widgets.py            Enhanced Tkinter widgets (legacy)
```

## config/ — Configuration

實際設定檔含 API token，**不進 git**（見 `.gitignore`），repo 只提供範本：

```
server.yaml.example            Odoo connection template (host, db, url)
token.yaml.example             API token template
server.yaml / server_*.yaml    [not tracked] actual connection settings
token.yaml / token_*.yaml      [not tracked] actual API tokens
```

## autolisp/ — AutoCAD Side

```
lisp/080_main.lsp              Entry point, module loader, OB:MCP-DISPATCH command
lisp/070_ob_mcp_dispatch.lsp   6 Odoo IPC actions, wraps mcp_dispatch.lsp
lisp/030_config.lsp            Path resolution, config directories
lisp/060_table_util.lsp        9-column TABLE read/write
lisp/050_block_util.lsp        Attribute block read/write
lisp/010_json_util.lsp         JSON parser (pure AutoLISP)
lisp/015_log_util.lsp          File logging (logs/autolisp_YYYY-MM-DD.log)
lisp/020_file_util.lsp         File I/O utilities
lisp/040_strip_mtext.lsp       MText formatting cleanup
build_vlx.lsp                  VLX build script (BUILD-ALL command)
dist/                          VLX output directory
docs/ARCHITECTURE.md           AutoLISP architecture doc
docs/IMPLEMENTATION_PLAN.md    Phase 6 implementation plan
docs/AUTOCAD_LT_RESEARCH.md    LT compatibility research
docs/INSTALL_GUIDE.md          AutoCAD installation manual
```

## libs/ — External

```
autocad-mcp/                   Git submodule (puran-water/autocad-mcp)
  src/autocad_mcp/             Python: File IPC backend, MCP server
  lisp-code/mcp_dispatch.lsp   AutoLISP: 70+ command dispatcher (IPC core)
  lisp-code/attribute_tools.lsp AutoLISP: Enhanced attribute operations
```

## doc/ — Operational Guides

```
DEPLOYMENT_GUIDE.md            Build & deploy instructions
CODE_SIGNING_GUIDE.md          VLX/EXE signing
ANTIVIRUS_SOLUTION.md          Antivirus whitelist guide
GEMINI_CLI_SETUP.md            Gemini CLI MCP configuration
MCP_INTEGRATION_PLAN.md        MCP integration planning
RELEASE_NOTES_v5.0.md          v5.0 release notes
README-DEVELOPMENT.md          Development environment setup
UI_IMPROVEMENT_PLAN.md         UI modernization plan
```

## Removed Files

已刪除的檔案（可由 git 歷史取回）：

```
mcp_server_fastmcp.py           → replaced by mcp_server_autocad.py
utility/util_mcp_sse_manager.py → replaced by util_mcp_manager.py
utility/util_mcp_sse_server.py  → replaced by mcp_server_autocad.py
ai_assistant/                   → TCP Socket + Named Pipe MCP，由 FastMCP 取代
forms/form_main.py              → 舊版 GUI，由 form_main_modern.py 取代
utility/util_com_server.py      → legacy COM server（呼叫者 autolisp/legacy/call_python.lsp）
utility/util_nlp_processor.py   → 未使用
utility/util_get_product_config.py → 未使用
```
