# -*- coding: utf-8 -*-
"""
Live COM integration tests — requires Full AutoCAD running with an open .dwg.

Run:  pytest tests/integration/test_com_live.py -m autocad -v
Skip: pytest tests/integration/ -m "not autocad"
"""
import pytest


pytestmark = pytest.mark.autocad


# ── Connection ───────────────────────────────────────────────

class TestCOMConnection:
    def test_connected(self, com_backend):
        assert com_backend.connected_autocad() is True

    def test_acad_not_none(self, com_backend):
        assert com_backend.acad is not None

    def test_doc_not_none(self, com_backend):
        assert com_backend.doc is not None


# ── get_active_layout ────────────────────────────────────────

class TestCOMGetActiveLayout:
    def test_returns_string(self, com_backend):
        result = com_backend.get_active_layout()
        assert result is None or isinstance(result, str), \
            f"Expected str, got {type(result).__name__}: {result}"

    def test_not_empty(self, com_backend):
        result = com_backend.get_active_layout()
        assert result, "Active layout should not be empty when a .dwg is open"


# ── get_doc_layouts ──────────────────────────────────────────

class TestCOMGetDocLayouts:
    def test_returns_list(self, com_backend):
        result = com_backend.get_doc_layouts()
        assert isinstance(result, list)

    def test_all_elements_are_strings(self, com_backend):
        result = com_backend.get_doc_layouts()
        for name in result:
            assert isinstance(name, str), \
                f"Expected str, got {type(name).__name__}: {name}"

    def test_model_excluded(self, com_backend):
        result = com_backend.get_doc_layouts()
        assert "Model" not in result

    def test_active_layout_in_list(self, com_backend):
        """The active layout should appear in get_doc_layouts (unless it's Model)"""
        active = com_backend.get_active_layout()
        layouts = com_backend.get_doc_layouts()
        if active and active != "Model":
            assert active in layouts


# ── get_layouts_values ───────────────────────────────────────

class TestCOMGetLayoutsValues:
    def test_returns_dict_with_all_key(self, com_backend):
        result = com_backend.get_layouts_values()
        assert isinstance(result, dict), f"Expected dict, got {type(result).__name__}"
        if result:
            assert "all" in result, f"Missing 'all' key, got keys: {list(result.keys())}"

    def test_all_is_list(self, com_backend):
        result = com_backend.get_layouts_values()
        if result and "all" in result:
            assert isinstance(result["all"], list)

    def test_layout_dicts_have_layout_name(self, com_backend):
        result = com_backend.get_layouts_values()
        if result and "all" in result:
            for item in result["all"]:
                assert "layout_name" in item, f"Missing 'layout_name' in: {list(item.keys())}"
                assert isinstance(item["layout_name"], str)


# ── get_single_layout_values ─────────────────────────────────

class TestCOMGetSingleLayoutValues:
    def test_returns_dict(self, com_backend):
        layouts = com_backend.get_doc_layouts()
        if not layouts:
            pytest.skip("No layouts in document")
        result = com_backend.get_single_layout_values(layouts[0])
        assert isinstance(result, dict)

    def test_has_layout_name_key(self, com_backend):
        layouts = com_backend.get_doc_layouts()
        if not layouts:
            pytest.skip("No layouts in document")
        result = com_backend.get_single_layout_values(layouts[0])
        if result:
            assert result.get("layout_name") == layouts[0]

    def test_nonexistent_layout_returns_empty(self, com_backend):
        result = com_backend.get_single_layout_values("__NONEXISTENT_LAYOUT__")
        assert result == {}


# ── get_block_attributes ─────────────────────────────────────

class TestCOMGetBlockAttributes:
    def test_returns_dict(self, com_backend):
        result = com_backend.get_block_attributes()
        assert isinstance(result, dict)

    def test_has_layout_name(self, com_backend):
        result = com_backend.get_block_attributes()
        if result:
            assert "layout_name" in result
            assert isinstance(result["layout_name"], str)


# ── get_layouts_header_id_to_pr ──────────────────────────────

class TestCOMGetLayoutsHeaderIdToPr:
    def test_returns_dict_with_all_key(self, com_backend):
        result = com_backend.get_layouts_header_id_to_pr()
        assert isinstance(result, dict)
        if result:
            assert "all" in result
            assert isinstance(result["all"], list)


# ── process_pr_no ────────────────────────────────────────────

class TestCOMProcessPrNo:
    def test_accepts_string_layout_name(self, com_backend):
        """The bug: process_pr_no used to crash with 'str has no attribute Block'.
        We only test the COM-side (layout resolution + block read), not the Odoo call."""
        from unittest.mock import MagicMock
        active = com_backend.get_active_layout()
        if not active:
            pytest.skip("No active layout")

        # Temporarily mock odoo_util so process_pr_no doesn't need Odoo
        saved = com_backend.odoo_util
        com_backend.odoo_util = MagicMock()
        com_backend.odoo_util.get_project.return_value = None
        try:
            # Must NOT raise "'str' object has no attribute 'Block'"
            com_backend.process_pr_no(active)
        finally:
            com_backend.odoo_util = saved

    def test_handles_nonexistent_layout(self, com_backend):
        # Must not crash
        com_backend.process_pr_no("__NONEXISTENT__")


# ── clear_table_id ───────────────────────────────────────────

class TestCOMClearTableId:
    def test_accepts_string(self, com_backend):
        layouts = com_backend.get_doc_layouts()
        if not layouts:
            pytest.skip("No layouts")
        # Must not crash with "'str' object has no attribute 'Block'"
        com_backend.clear_table_id(layouts[0])

    def test_accepts_none(self, com_backend):
        # Uses active layout
        com_backend.clear_table_id()


# ── Drawing operations (read-only safety) ────────────────────

class TestCOMListLayers:
    def test_returns_dict(self, com_backend):
        result = com_backend.list_layers()
        assert isinstance(result, dict)
        assert "layers" in result
        assert isinstance(result["layers"], list)

    def test_layers_have_name(self, com_backend):
        result = com_backend.list_layers()
        for layer in result["layers"]:
            assert "name" in layer
            assert isinstance(layer["name"], str)


class TestCOMScanElements:
    def test_raises_not_implemented(self, com_backend):
        """COM 端未實作 —— 不可回傳空結果讓呼叫端誤以為圖面沒有元素"""
        import pytest as _pytest
        with _pytest.raises(NotImplementedError, match="IPC"):
            com_backend.scan_elements()
