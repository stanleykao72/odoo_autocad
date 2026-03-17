# -*- coding: utf-8 -*-
"""
Tests for IPC backend (UtilAutoCADIPC).

Verifies:
- get_active_layout() returns str
- get_doc_layouts() returns List[str]
- get_layouts_values() returns {"all": [...]} — not bare list
- get_layouts_header_id_to_pr() returns {"all": [...]}
- get_block_attributes() returns dict
- get_single_layout_values() returns dict
- set_block_attributes() returns True on success
"""
import pytest
from unittest.mock import MagicMock, AsyncMock, patch
from types import SimpleNamespace


@pytest.fixture
def ipc_backend(mock_log_util):
    from utility.util_autocad_ipc import UtilAutoCADIPC
    util = UtilAutoCADIPC(log_util=mock_log_util)
    return util


def _make_result(ok=True, payload=None, error=None):
    """Create a CommandResult-like object"""
    return SimpleNamespace(ok=ok, payload=payload, error=error)


def _patch_dispatch(ipc_backend, responses):
    """Patch _dispatch to return predefined responses by command name.

    Args:
        responses: dict of {command: CommandResult} or a single CommandResult for all commands
    """
    if isinstance(responses, dict):
        async def mock_dispatch(command, params=None):
            return responses.get(command, _make_result(ok=False, error="Unknown command"))
    else:
        async def mock_dispatch(command, params=None):
            return responses

    ipc_backend._dispatch = mock_dispatch
    ipc_backend._run_async = lambda coro: __import__('asyncio').get_event_loop().run_until_complete(coro) \
        if hasattr(coro, '__await__') else coro

    # Simpler: just mock _run_async to call the coro
    import asyncio
    original_run_async = ipc_backend._run_async

    def sync_run(coro):
        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(coro)
        finally:
            loop.close()

    ipc_backend._run_async = sync_run


class TestIPCGetActiveLayout:
    def test_returns_string(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "drawing-info": _make_result(ok=True, payload={
                "active_layout": "Layout1",
                "file_name": "test.dwg"
            })
        })

        result = ipc_backend.get_active_layout()
        assert isinstance(result, str)
        assert result == "Layout1"

    def test_returns_none_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "drawing-info": _make_result(ok=False, error="no doc")
        })

        result = ipc_backend.get_active_layout()
        assert result is None


class TestIPCGetDocLayouts:
    def test_returns_list_of_strings(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_get_layouts": _make_result(ok=True, payload={
                "layouts": ["Layout1", "Layout2", "Model"]
            })
        })

        result = ipc_backend.get_doc_layouts()
        assert isinstance(result, list)
        assert all(isinstance(name, str) for name in result)
        assert "Model" not in result
        assert result == ["Layout1", "Layout2"]

    def test_returns_empty_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_get_layouts": _make_result(ok=False, error="timeout")
        })

        result = ipc_backend.get_doc_layouts()
        assert result == []


class TestIPCGetLayoutsValues:
    def test_returns_dict_with_all_key_from_list_payload(self, ipc_backend):
        """When AutoLISP returns a list, it must be wrapped in {"all": [...]}"""
        _patch_dispatch(ipc_backend, {
            "odoo_extract_tables": _make_result(ok=True, payload=[
                {"layout_name": "L1", "header_id": "H1", "detail": []},
                {"layout_name": "L2", "header_id": "H2", "detail": []},
            ])
        })

        result = ipc_backend.get_layouts_values()
        assert isinstance(result, dict)
        assert "all" in result
        assert isinstance(result["all"], list)
        assert len(result["all"]) == 2

    def test_returns_dict_with_all_key_already_wrapped(self, ipc_backend):
        """When payload already has {"all": [...]}, pass through"""
        _patch_dispatch(ipc_backend, {
            "odoo_extract_tables": _make_result(ok=True, payload={
                "all": [{"layout_name": "L1"}]
            })
        })

        result = ipc_backend.get_layouts_values()
        assert isinstance(result, dict)
        assert "all" in result

    def test_returns_empty_dict_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_extract_tables": _make_result(ok=False, error="timeout")
        })

        result = ipc_backend.get_layouts_values()
        assert isinstance(result, dict)
        # Must not return [] (bare list)
        assert not isinstance(result, list)


class TestIPCGetSingleLayoutValues:
    def test_returns_dict(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_extract_table_for_layout": _make_result(ok=True, payload={
                "layout_name": "Layout1",
                "header_id": "H001",
                "detail": [{"product_no": "PROD-A"}]
            })
        })

        result = ipc_backend.get_single_layout_values("Layout1")
        assert isinstance(result, dict)
        assert result.get("layout_name") == "Layout1"

    def test_returns_empty_dict_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_extract_table_for_layout": _make_result(ok=False, error="timeout")
        })

        result = ipc_backend.get_single_layout_values("X")
        assert result == {}


class TestIPCGetBlockAttributes:
    def test_returns_dict(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_get_block_attrs": _make_result(ok=True, payload={
                "pr_no": "PR001",
                "project_name": "TestProject",
                "layout_name": "Layout1"
            })
        })

        result = ipc_backend.get_block_attributes()
        assert isinstance(result, dict)
        assert result["pr_no"] == "PR001"

    def test_returns_empty_dict_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_get_block_attrs": _make_result(ok=False, error="Unknown command")
        })

        result = ipc_backend.get_block_attributes()
        assert result == {}


class TestIPCSetBlockAttributes:
    def test_returns_true_on_success(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_set_block_attrs": _make_result(ok=True)
        })

        result = ipc_backend.set_block_attributes({"pr_no": "PR002"}, "Layout1")
        assert result is True

    def test_raises_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_set_block_attrs": _make_result(ok=False, error="write failed")
        })

        with pytest.raises(RuntimeError):
            ipc_backend.set_block_attributes({"pr_no": "X"})


class TestIPCGetLayoutsHeaderIdToPr:
    def test_returns_dict_with_all_key(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_get_header_ids": _make_result(ok=True, payload={
                "header_ids": ["H001", "H002"]
            })
        })

        result = ipc_backend.get_layouts_header_id_to_pr()
        assert isinstance(result, dict)
        assert "all" in result
        assert result["all"] == ["H001", "H002"]

    def test_returns_empty_dict_on_failure(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "odoo_get_header_ids": _make_result(ok=False, error="timeout")
        })

        result = ipc_backend.get_layouts_header_id_to_pr()
        assert result == {}


class TestIPCDrawOperations:
    def test_draw_line(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "create-line": _make_result(ok=True)
        })
        result = ipc_backend.draw_line([0, 0], [10, 10])
        assert result is True

    def test_draw_circle(self, ipc_backend):
        _patch_dispatch(ipc_backend, {
            "create-circle": _make_result(ok=True)
        })
        result = ipc_backend.draw_circle([5, 5], 10)
        assert result is True
