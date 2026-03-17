# -*- coding: utf-8 -*-
"""
AutoCAD IPC 後端 — 透過 autocad-mcp File IPC 驅動 AutoCAD LT 2024+

提供與 UtilAutoCAD (COM) 相似的介面，但使用 File IPC 通訊。
"""

import asyncio
import json
import logging
import sys
import os

# Add autocad-mcp to path
_autocad_mcp_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "libs", "autocad-mcp", "src"
)
if _autocad_mcp_path not in sys.path:
    sys.path.insert(0, _autocad_mcp_path)

_logger = logging.getLogger(__name__)


class UtilAutoCADIPC:
    """AutoCAD IPC 後端 — 使用 autocad-mcp File IPC"""

    def __init__(self, log_util=None, target_hwnd=None):
        self.log = log_util
        self._backend = None
        self._initialized = False
        self._loop = None
        self._target_hwnd = target_hwnd  # specific AutoCAD window HWND
        self.acad = None  # Compatibility: None means not COM-connected
        self.project_id = None
        self.project_name = None
        self.pr_no = None
        self.job_working_plan_id = None
        self.job_working_plan_name = None

    def _log(self, msg):
        if self.log:
            self.log.safe_log_insert(msg)
        else:
            _logger.info(msg.rstrip('\n'))

    def _get_loop(self):
        """Get or create an event loop for running async code"""
        if self._loop is None or self._loop.is_closed():
            self._loop = asyncio.new_event_loop()
        return self._loop

    def _run_async(self, coro):
        """Run an async coroutine synchronously"""
        loop = self._get_loop()
        return loop.run_until_complete(coro)

    async def _get_backend(self):
        """Lazy-init the File IPC backend"""
        if self._backend is None:
            try:
                from autocad_mcp.backends.file_ipc import FileIPCBackend
                self._backend = FileIPCBackend()
                # Override HWND if a specific target was requested
                if self._target_hwnd:
                    self._backend._hwnd = self._target_hwnd
                    self._backend._command_hwnd = self._backend._find_command_line_hwnd()
                    self._log(f"[IPC] Using target HWND: {self._target_hwnd}\n")
                init_result = await self._backend.initialize()
                if hasattr(init_result, 'ok') and init_result.ok:
                    self._initialized = True
                    self._log("[IPC] File IPC backend initialized (ping OK)\n")
                else:
                    self._initialized = False
                    error = getattr(init_result, 'error', 'unknown')
                    self._log(f"[IPC] File IPC backend initialized (ping failed: {error})\n")
                # Monkey-patch _dispatch_unlocked to log JSON parse errors
                self._patch_dispatch_logging()
            except Exception as e:
                self._log(f"[IPC] Failed to initialize File IPC backend: {e}\n")
                raise
        return self._backend

    def _patch_dispatch_logging(self):
        """Patch FileIPCBackend to log JSON parse errors instead of silent pass"""
        import json as _json
        import asyncio as _asyncio
        import time as _time
        import uuid as _uuid
        backend = self._backend
        log = self._log

        def _fix_json_backslashes(text):
            """Fix unescaped backslashes in JSON from AutoLISP serializer."""
            VALID = set('"\\' + '/bfnrtu')
            out = []
            i = 0
            while i < len(text):
                ch = text[i]
                if ch == '\\':
                    if i + 1 < len(text) and text[i + 1] in VALID:
                        out.append(ch)
                        out.append(text[i + 1])
                        i += 2
                    else:
                        out.append('\\\\')
                        i += 1
                else:
                    out.append(ch)
                    i += 1
            return ''.join(out)

        original_dispatch = backend._dispatch_unlocked

        async def patched_dispatch(command, params):
            """Wrapped _dispatch_unlocked with JSON error logging"""
            from autocad_mcp.backends.file_ipc import TIMEOUT, POLL_INTERVAL, CommandResult
            import pathlib

            request_id = _uuid.uuid4().hex[:12]
            ipc_dir = backend._ipc_dir
            cmd_file = ipc_dir / f"autocad_mcp_cmd_{request_id}.json"
            result_file = ipc_dir / f"autocad_mcp_result_{request_id}.json"
            tmp_file = cmd_file.with_suffix(".tmp")

            try:
                clean_params = {k: v for k, v in params.items() if v is not None}
                payload = {
                    "request_id": request_id,
                    "command": command,
                    "params": clean_params,
                    "ts": _time.time(),
                }
                json_str = _json.dumps(payload, ensure_ascii=False)
                try:
                    tmp_file.write_text(json_str, encoding="cp950")
                except UnicodeEncodeError:
                    tmp_file.write_text(json_str, encoding="utf-8")
                tmp_file.rename(cmd_file)

                backend._type_dispatch_trigger()

                deadline = _time.time() + TIMEOUT
                parse_error_logged = False
                while _time.time() < deadline:
                    if result_file.exists():
                        try:
                            # AutoCAD writes in system codepage (cp950 for Chinese).
                            # Try cp950 FIRST to avoid Big5 0x5C phantom backslash,
                            # then UTF-8, then cp1252.
                            text = None
                            used_enc = None
                            for enc in ("cp950", "utf-8", "cp1252"):
                                try:
                                    text = result_file.read_text(encoding=enc)
                                    used_enc = enc
                                    break
                                except (UnicodeDecodeError, ValueError):
                                    continue
                            if text is None:
                                text = result_file.read_text(
                                    encoding="utf-8", errors="replace")
                                used_enc = "utf-8(replace)"

                            data = _json.loads(text)
                            if data.get("request_id") == request_id:
                                return CommandResult(
                                    ok=data.get("ok", False),
                                    payload=data.get("payload"),
                                    error=data.get("error"),
                                )
                        except _json.JSONDecodeError as e:
                            # Fix unescaped backslashes from AutoLISP JSON
                            try:
                                fixed = _fix_json_backslashes(text)
                                data = _json.loads(fixed)
                                if data.get("request_id") == request_id:
                                    if not parse_error_logged:
                                        log(f"[IPC] Fixed JSON backslash ({used_enc}) in {command}\n")
                                    return CommandResult(
                                        ok=data.get("ok", False),
                                        payload=data.get("payload"),
                                        error=data.get("error"),
                                    )
                            except _json.JSONDecodeError:
                                pass
                            if not parse_error_logged:
                                fsize = result_file.stat().st_size
                                log(f"[IPC] DIAG: JSON error ({used_enc}): {e}\n")
                                log(f"[IPC] DIAG: size={fsize}, first 300: {text[:300]}\n")
                                log(f"[IPC] DIAG: last 200: {text[-200:]}\n")
                                parse_error_logged = True
                        except OSError as e:
                            if not parse_error_logged:
                                log(f"[IPC] DIAG: OS error: {e}\n")
                                parse_error_logged = True
                    await _asyncio.sleep(POLL_INTERVAL)

                return CommandResult(ok=False, error=f"Timeout waiting for result (request_id={request_id})")
            finally:
                for f in (cmd_file, result_file, tmp_file):
                    try:
                        f.unlink(missing_ok=True)
                    except OSError:
                        pass

        backend._dispatch_unlocked = patched_dispatch

    # Commands that may take longer than the default IPC timeout (10s)
    _SLOW_COMMANDS = {"odoo_extract_tables", "odoo_extract_table_for_layout",
                      "odoo_get_header_ids", "odoo_write_ids"}
    _SLOW_TIMEOUT = 60.0  # seconds

    async def _dispatch(self, command, params=None):
        """Send a command via File IPC and return the result"""
        backend = await self._get_backend()
        # For slow commands, temporarily increase the backend timeout
        if command in self._SLOW_COMMANDS:
            import autocad_mcp.backends.file_ipc as _fipc
            saved_timeout = _fipc.TIMEOUT
            _fipc.TIMEOUT = self._SLOW_TIMEOUT
            try:
                result = await backend._dispatch(command, params or {})
            finally:
                _fipc.TIMEOUT = saved_timeout
        else:
            result = await backend._dispatch(command, params or {})
        payload_preview = str(result.payload)[:200] if result.payload else None
        self._log(f"[IPC] dispatch({command}) -> ok={result.ok}, error={result.error}, payload={payload_preview}\n")

        # Diagnostic: if timeout, check for orphaned result file (truncated JSON)
        if not result.ok and result.error and "Timeout" in str(result.error):
            self._diagnose_timeout(command, result)

        return result

    def _diagnose_timeout(self, command, result):
        """Diagnose timeout failures — check for truncated response files"""
        import pathlib
        ipc_dir = pathlib.Path("C:/temp")
        try:
            # Look for any leftover result files
            result_files = list(ipc_dir.glob("autocad_mcp_result_*.json"))
            tmp_files = list(ipc_dir.glob("autocad_mcp_result_*.json.tmp"))
            self._log(f"[IPC] DIAG: {command} timeout. "
                      f"Orphaned result files: {len(result_files)}, "
                      f"tmp files: {len(tmp_files)}\n")
            for rf in result_files[:3]:  # Check up to 3 files
                try:
                    size = rf.stat().st_size
                    preview = rf.read_text(encoding="utf-8", errors="replace")[:300]
                    self._log(f"[IPC] DIAG: {rf.name} ({size} bytes): {preview}...\n")
                except Exception as e:
                    self._log(f"[IPC] DIAG: {rf.name} read error: {e}\n")
            for tf in tmp_files[:3]:
                try:
                    size = tf.stat().st_size
                    self._log(f"[IPC] DIAG: {tf.name} ({size} bytes)\n")
                except Exception as e:
                    self._log(f"[IPC] DIAG: {tf.name} error: {e}\n")
        except Exception as e:
            self._log(f"[IPC] DIAG error: {e}\n")

    # === Connection & Status ===

    def connected_autocad(self):
        """Check if IPC connection to AutoCAD is available"""
        try:
            result = self._run_async(self._dispatch("ping"))
            return result.ok if hasattr(result, 'ok') else bool(result)
        except Exception:
            return False

    def connect_autocad(self, main_body=None):
        """Establish IPC connection to AutoCAD"""
        self._log("[IPC] Connecting to AutoCAD via File IPC...\n")
        try:
            self._run_async(self._get_backend())
            # initialize() already did a ping - use that result
            if getattr(self, '_initialized', False):
                self._log("[IPC] AutoCAD IPC connection established\n")
                self._check_odoo_extensions()
                return
            # initialize() ping failed - retry with delay
            import time
            for attempt in range(3):
                time.sleep(1.0)
                self._log(f"[IPC] Retry ping ({attempt + 1}/3)...\n")
                if self.connected_autocad():
                    self._log("[IPC] AutoCAD IPC connection established\n")
                    return
            self._log("[IPC] AutoCAD not responding to IPC ping\n")
            self._log("[IPC] Ensure mcp_dispatch.lsp or McpDispatch.vlx is loaded in AutoCAD\n")
        except Exception as e:
            self._log(f"[IPC] Connection failed: {e}\n")

    def _check_odoo_extensions(self):
        """Check if Odoo extensions (070_ob_mcp_dispatch.lsp) are loaded in AutoCAD"""
        try:
            result = self._run_async(self._dispatch("odoo_get_block_attrs"))
            if hasattr(result, 'error') and 'Unknown command' in str(result.error or ''):
                self._log("[IPC] WARNING: Odoo extensions NOT loaded in AutoCAD!\n")
                self._log("[IPC] Please run in AutoCAD: (load \"080_main.lsp\")\n")
            else:
                self._log("[IPC] Odoo extensions loaded (6 actions available)\n")
        except Exception:
            pass

    # === Layout Management ===

    def get_active_layout(self):
        """Get the current active layout name"""
        try:
            result = self._run_async(self._dispatch("drawing-info"))
            if hasattr(result, 'ok') and result.ok and result.payload:
                payload = result.payload
                # payload may be dict or JSON string
                if isinstance(payload, str):
                    import json
                    payload = json.loads(payload)
                layout = payload.get('active_layout')
                self._log(f"[IPC] active_layout: {layout}\n")
                return layout
            self._log(f"[IPC] get_active_layout: no payload (ok={getattr(result, 'ok', '?')})\n")
            return None
        except Exception as e:
            self._log(f"[IPC] get_active_layout failed: {e}\n")
            return None

    def get_doc_layouts(self):
        """Get list of layout names (excluding Model) via Odoo extension"""
        try:
            result = self._run_async(self._dispatch("odoo_get_layouts"))
            if hasattr(result, 'ok') and result.ok and result.payload:
                payload = result.payload
                if isinstance(payload, str):
                    import json as _json
                    payload = _json.loads(payload)
                layouts = payload.get('layouts', []) if isinstance(payload, dict) else payload
                return [l for l in layouts if l != 'Model']
            self._log(f"[IPC] get_doc_layouts: no payload (error={getattr(result, 'error', '?')})\n")
            return []
        except Exception as e:
            self._log(f"[IPC] get_doc_layouts failed: {e}\n")
            return []

    # === Odoo-specific Operations (via ob_mcp_dispatch.lsp) ===

    def get_layouts_values(self):
        """Extract TABLE + Block data from all layouts (Odoo action) — fallback"""
        try:
            result = self._run_async(self._dispatch("odoo_extract_tables"))
            if hasattr(result, 'ok') and result.ok:
                return result.payload if result.payload else []
            self._log(f"[IPC] get_layouts_values failed: {getattr(result, 'error', 'unknown')}\n")
            return []
        except Exception as e:
            self._log(f"[IPC] get_layouts_values error: {e}\n")
            return []

    def get_single_layout_values(self, layout_name):
        """Extract TABLE + Block data from a single layout by name"""
        try:
            result = self._run_async(self._dispatch(
                "odoo_extract_table_for_layout",
                {"layout_name": layout_name}
            ))
            if hasattr(result, 'ok') and result.ok:
                return result.payload if result.payload else {}
            self._log(f"[IPC] get_single_layout_values({layout_name}) failed: {getattr(result, 'error', 'unknown')}\n")
            return {}
        except Exception as e:
            self._log(f"[IPC] get_single_layout_values({layout_name}) error: {e}\n")
            return {}

    def get_layouts_header_id_to_pr(self):
        """Collect all header_ids from TABLEs (Odoo action)
        Returns {"all": [header_id1, header_id2, ...]} to match COM mode format."""
        try:
            result = self._run_async(self._dispatch("odoo_get_header_ids"))
            if hasattr(result, 'ok') and result.ok and result.payload:
                payload = result.payload
                # Extract header_ids list and wrap in {"all": [...]}
                if isinstance(payload, dict) and 'header_ids' in payload:
                    ids = payload['header_ids']
                    self._log(f"[IPC] header_ids: {ids}\n")
                    return {"all": ids}
                return payload if isinstance(payload, dict) else {}
            return {}
        except Exception as e:
            self._log(f"[IPC] get_layouts_header_id_to_pr error: {e}\n")
            return {}

    def set_layouts_tables_id(self, boq_list):
        """Write header_id + detail_id back to TABLEs (Odoo action)"""
        try:
            # boq_list is the response from Odoo import2boq
            payload = {"all": boq_list} if isinstance(boq_list, list) else boq_list
            result = self._run_async(self._dispatch("odoo_write_ids", payload))
            if hasattr(result, 'ok') and result.ok:
                self._log("[IPC] IDs written back to TABLEs\n")
            else:
                self._log(f"[IPC] ID writeback failed: {getattr(result, 'error', 'unknown')}\n")
        except Exception as e:
            self._log(f"[IPC] set_layouts_tables_id error: {e}\n")

    def get_block_attributes(self):
        """Read attribute block values from current layout (Odoo action)"""
        try:
            result = self._run_async(self._dispatch("odoo_get_block_attrs"))
            if hasattr(result, 'ok') and result.ok:
                return result.payload if result.payload else {}
            error = getattr(result, 'error', 'unknown')
            self._log(f"[IPC] get_block_attributes: {error}\n")
            return {}
        except Exception as e:
            self._log(f"[IPC] get_block_attributes error: {e}\n")
            return {}

    def set_block_attributes(self, attrs, layout_name=None):
        """Write attributes to Block (Odoo action). Returns True on success."""
        try:
            params = dict(attrs) if not isinstance(attrs, dict) else attrs
            if layout_name:
                params["layout_name"] = layout_name
            result = self._run_async(self._dispatch("odoo_set_block_attrs", params))
            if hasattr(result, 'ok') and result.ok:
                self._log("[IPC] Block attributes written\n")
                return True
            else:
                error = getattr(result, 'error', 'unknown')
                self._log(f"[IPC] set_block_attributes failed: {error}\n")
                raise RuntimeError(f"寫入屬性失敗: {error}")
        except RuntimeError:
            raise
        except Exception as e:
            self._log(f"[IPC] set_block_attributes error: {e}\n")
            raise

    def clear_table_id(self, layout=None):
        """Clear TABLE IDs"""
        try:
            result = self._run_async(self._dispatch("odoo_clear_ids"))
            if hasattr(result, 'ok') and result.ok:
                self._log("[IPC] TABLE IDs cleared\n")
        except Exception as e:
            self._log(f"[IPC] clear_table_id error: {e}\n")

    def clear_all_tables_id(self):
        """Clear TABLE IDs in all layouts"""
        self.clear_table_id()

    # === Drawing Operations (via autocad-mcp built-in commands) ===

    def draw_line(self, start_point, end_point, layer="0"):
        """Draw a line via IPC"""
        try:
            result = self._run_async(self._dispatch("create-line", {
                "x1": start_point[0], "y1": start_point[1],
                "x2": end_point[0], "y2": end_point[1],
                "layer": layer
            }))
            return hasattr(result, 'ok') and result.ok
        except Exception:
            return False

    def draw_circle(self, center_point, radius, layer="0"):
        """Draw a circle via IPC"""
        try:
            result = self._run_async(self._dispatch("create-circle", {
                "cx": center_point[0], "cy": center_point[1],
                "radius": radius, "layer": layer
            }))
            return hasattr(result, 'ok') and result.ok
        except Exception:
            return False

    def set_layer(self, layer_name, color=7, create_if_not_exist=True):
        """Set or create a layer via IPC"""
        try:
            if create_if_not_exist:
                self._run_async(self._dispatch("layer-create", {
                    "name": layer_name, "color": color
                }))
            self._run_async(self._dispatch("layer-set-current", {
                "name": layer_name
            }))
        except Exception as e:
            self._log(f"[IPC] set_layer error: {e}\n")

    def list_layers(self, filter_type="all", sort_by="name", include_details=True):
        """List layers via IPC"""
        try:
            result = self._run_async(self._dispatch("layer-list"))
            if hasattr(result, 'ok') and result.ok:
                return {"success": True, "layers": result.payload}
            return {"success": False, "layers": []}
        except Exception:
            return {"success": False, "layers": []}

    def scan_elements(self, element_type="all", **kwargs):
        """Scan drawing elements via IPC"""
        try:
            result = self._run_async(self._dispatch("entity-list", {
                "type": element_type
            }))
            if hasattr(result, 'ok') and result.ok:
                return {"success": True, "elements": result.payload}
            return {"success": False, "elements": []}
        except Exception:
            return {"success": False, "elements": []}
