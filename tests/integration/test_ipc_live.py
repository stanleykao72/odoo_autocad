# -*- coding: utf-8 -*-
"""
Live IPC integration tests — requires AutoCAD (LT or Full) with McpDispatch.vlx loaded.

Run:  pytest tests/integration/test_ipc_live.py -m ipc -v
Skip: pytest tests/integration/ -m "not ipc"
"""
import pytest


pytestmark = pytest.mark.ipc


# ── Connection ───────────────────────────────────────────────

class TestIPCConnection:
    def test_connected(self, ipc_backend):
        assert ipc_backend.connected_autocad() is True


# ── get_active_layout ────────────────────────────────────────

class TestIPCGetActiveLayout:
    def test_returns_string_or_none(self, ipc_backend):
        """May return None if drawing-info doesn't provide active_layout"""
        result = ipc_backend.get_active_layout()
        assert result is None or isinstance(result, str), \
            f"Expected str or None, got {type(result).__name__}: {result}"


# ── get_doc_layouts ──────────────────────────────────────────

class TestIPCGetDocLayouts:
    def test_returns_list(self, ipc_backend):
        result = ipc_backend.get_doc_layouts()
        assert isinstance(result, list)

    def test_all_elements_are_strings(self, ipc_backend):
        result = ipc_backend.get_doc_layouts()
        for name in result:
            assert isinstance(name, str), \
                f"Expected str, got {type(name).__name__}: {name}"

    def test_model_excluded(self, ipc_backend):
        result = ipc_backend.get_doc_layouts()
        assert "Model" not in result


# ── get_layouts_values ───────────────────────────────────────

class TestIPCGetLayoutsValues:
    def test_returns_dict_with_all_key(self, ipc_backend):
        result = ipc_backend.get_layouts_values()
        assert isinstance(result, dict), f"Expected dict, got {type(result).__name__}"
        if result:
            assert "all" in result, f"Missing 'all' key, got keys: {list(result.keys())}"

    def test_all_is_list(self, ipc_backend):
        result = ipc_backend.get_layouts_values()
        if result and "all" in result:
            assert isinstance(result["all"], list)


# ── get_single_layout_values ─────────────────────────────────

class TestIPCGetSingleLayoutValues:
    def test_returns_dict(self, ipc_backend):
        layouts = ipc_backend.get_doc_layouts()
        if not layouts:
            pytest.skip("No layouts (odoo_get_layouts not loaded?)")
        result = ipc_backend.get_single_layout_values(layouts[0])
        assert isinstance(result, dict)

    def test_nonexistent_layout_returns_empty(self, ipc_backend):
        result = ipc_backend.get_single_layout_values("__NONEXISTENT_LAYOUT__")
        assert result == {}


# ── get_block_attributes ─────────────────────────────────────

class TestIPCGetBlockAttributes:
    def test_returns_dict(self, ipc_backend):
        result = ipc_backend.get_block_attributes()
        assert isinstance(result, dict)


# ── get_layouts_header_id_to_pr ──────────────────────────────

class TestIPCGetLayoutsHeaderIdToPr:
    def test_returns_dict_with_all_key(self, ipc_backend):
        result = ipc_backend.get_layouts_header_id_to_pr()
        assert isinstance(result, dict)
        if result:
            assert "all" in result
            assert isinstance(result["all"], list)


# ── Drawing operations ───────────────────────────────────────

class TestIPCListLayers:
    def test_returns_dict(self, ipc_backend):
        result = ipc_backend.list_layers()
        assert isinstance(result, dict)
        assert "layers" in result or "success" in result

class TestIPCScanElements:
    def test_returns_dict(self, ipc_backend):
        result = ipc_backend.scan_elements()
        assert isinstance(result, dict)


# ── Format consistency with COM ──────────────────────────────

class TestIPCFormatConsistency:
    """Verify IPC returns the same format as COM — the ABC contract."""

    def test_get_layouts_values_never_returns_bare_list(self, ipc_backend):
        """The original bug: IPC returned [] instead of {"all": []}"""
        result = ipc_backend.get_layouts_values()
        assert not isinstance(result, list), \
            "get_layouts_values must return dict, not list"

    def test_get_layouts_header_id_to_pr_never_returns_bare_list(self, ipc_backend):
        result = ipc_backend.get_layouts_header_id_to_pr()
        assert not isinstance(result, list), \
            "get_layouts_header_id_to_pr must return dict, not list"
