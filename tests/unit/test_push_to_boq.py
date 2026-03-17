# -*- coding: utf-8 -*-
"""
Tests for UtilPushToBoq — mode-agnostic business logic.

Verifies:
- No mode checks in push_to_boq (no 'is_ipc', no 'mode == ipc')
- Per-layout path works with any backend
- Fallback to bulk path works
- Progress callback is called
- {"all": [...]} wrapping for Odoo API
- Error handling
"""
import pytest
import inspect
from unittest.mock import MagicMock, call


@pytest.fixture
def mock_autocad():
    mock = MagicMock()
    mock.pr_no = "PR001"
    return mock


@pytest.fixture
def mock_odoo():
    return MagicMock()


@pytest.fixture
def push_util(mock_odoo, mock_autocad, mock_log_util):
    from utility.util_push_to_boq import UtilPushToBoq
    return UtilPushToBoq(mock_odoo, mock_autocad, mock_log_util)


class TestPushToBoqNoModeCheck:
    def test_no_mode_check_in_source(self):
        """push_to_boq.py must not contain mode-specific checks"""
        from utility.util_push_to_boq import UtilPushToBoq
        source = inspect.getsource(UtilPushToBoq)
        assert "is_ipc" not in source
        assert "mode == 'ipc'" not in source
        assert "mode == \"ipc\"" not in source
        assert "mode == 'com'" not in source


class TestPushToBoqPerLayoutPath:
    """Per-layout extraction — both COM and IPC should work the same"""

    def test_per_layout_success(self, push_util, mock_autocad, mock_odoo):
        """When get_doc_layouts returns names, per-layout extraction runs"""
        mock_autocad.get_doc_layouts.return_value = ["L1", "L2"]
        mock_autocad.get_single_layout_values.side_effect = [
            {"layout_name": "L1", "pr_no": "PR001", "header_id": "H1", "detail": [{"product_no": "A"}]},
            {"layout_name": "L2", "pr_no": "PR001", "header_id": "H2", "detail": [{"product_no": "B"}]},
        ]
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}, {"layout_name": "L2"}]

        result = push_util.push_to_boq()

        assert result is True
        # import2boq should receive {"all": [...]}
        call_arg = mock_odoo.import2boq.call_args[0][0]
        assert isinstance(call_arg, dict)
        assert "all" in call_arg
        assert len(call_arg["all"]) == 2

    def test_per_layout_injects_pr_no(self, push_util, mock_autocad, mock_odoo):
        """pr_no is injected if missing from layout data"""
        mock_autocad.get_doc_layouts.return_value = ["L1"]
        mock_autocad.get_single_layout_values.return_value = {
            "layout_name": "L1", "header_id": "H1", "detail": []
        }  # no pr_no
        mock_autocad.pr_no = "PR999"
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}]

        push_util.push_to_boq()

        call_arg = mock_odoo.import2boq.call_args[0][0]
        assert call_arg["all"][0]["pr_no"] == "PR999"

    def test_per_layout_skips_empty_data(self, push_util, mock_autocad, mock_odoo):
        """Empty layout data is skipped"""
        mock_autocad.get_doc_layouts.return_value = ["L1", "L2"]
        mock_autocad.get_single_layout_values.side_effect = [
            {"layout_name": "L1", "header_id": "H1", "detail": []},
            {},  # L2 is empty
        ]
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}]

        push_util.push_to_boq()

        call_arg = mock_odoo.import2boq.call_args[0][0]
        assert len(call_arg["all"]) == 1


class TestPushToBoqFallbackPath:
    """Bulk extraction fallback — when get_doc_layouts returns empty"""

    def test_fallback_when_no_layouts(self, push_util, mock_autocad, mock_odoo):
        """Falls back to get_layouts_values when get_doc_layouts is empty"""
        mock_autocad.get_doc_layouts.return_value = []
        mock_autocad.get_layouts_values.return_value = {
            "all": [{"layout_name": "L1", "header_id": "H1", "detail": []}]
        }
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}]

        result = push_util.push_to_boq()

        assert result is True
        mock_autocad.get_layouts_values.assert_called_once()

    def test_fallback_handles_already_wrapped(self, push_util, mock_autocad, mock_odoo):
        """get_layouts_values returning {"all": [...]} is handled correctly"""
        mock_autocad.get_doc_layouts.return_value = []
        mock_autocad.get_layouts_values.return_value = {"all": [{"layout_name": "L1"}]}
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}]

        push_util.push_to_boq()

        call_arg = mock_odoo.import2boq.call_args[0][0]
        assert "all" in call_arg

    def test_fallback_per_layout_all_empty_then_bulk(self, push_util, mock_autocad, mock_odoo):
        """If per-layout all returns empty, fallback to bulk"""
        mock_autocad.get_doc_layouts.return_value = ["L1"]
        mock_autocad.get_single_layout_values.return_value = {}  # empty
        mock_autocad.get_layouts_values.return_value = {
            "all": [{"layout_name": "L1", "header_id": "H1", "detail": []}]
        }
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}]

        result = push_util.push_to_boq()
        assert result is True
        mock_autocad.get_layouts_values.assert_called_once()


class TestPushToBoqErrorHandling:
    def test_returns_false_when_no_data(self, push_util, mock_autocad, mock_odoo):
        mock_autocad.get_doc_layouts.return_value = []
        mock_autocad.get_layouts_values.return_value = {}

        result = push_util.push_to_boq()
        assert result is False

    def test_raises_on_odoo_error_string(self, push_util, mock_autocad, mock_odoo):
        mock_autocad.get_doc_layouts.return_value = ["L1"]
        mock_autocad.get_single_layout_values.return_value = {"layout_name": "L1", "detail": []}
        mock_odoo.import2boq.return_value = "Error: invalid data"

        with pytest.raises(RuntimeError, match="Error: invalid data"):
            push_util.push_to_boq()

    def test_returns_false_when_odoo_returns_empty(self, push_util, mock_autocad, mock_odoo):
        mock_autocad.get_doc_layouts.return_value = ["L1"]
        mock_autocad.get_single_layout_values.return_value = {"layout_name": "L1", "detail": []}
        mock_odoo.import2boq.return_value = []

        result = push_util.push_to_boq()
        assert result is False


class TestPushToBoqProgressCallback:
    def test_progress_callback_called(self, push_util, mock_autocad, mock_odoo):
        mock_autocad.get_doc_layouts.return_value = ["L1"]
        mock_autocad.get_single_layout_values.return_value = {"layout_name": "L1", "detail": []}
        mock_odoo.import2boq.return_value = [{"layout_name": "L1"}]

        callback = MagicMock()
        callback._cancelled = False

        push_util.push_to_boq(progress_callback=callback)

        assert callback.call_count >= 3  # at least start, extract, done

    def test_cancellation_stops_processing(self, push_util, mock_autocad, mock_odoo):
        mock_autocad.get_doc_layouts.return_value = ["L1", "L2", "L3"]

        callback = MagicMock()
        callback._cancelled = True  # cancelled from the start

        result = push_util.push_to_boq(progress_callback=callback)
        assert result is False
        # Should not have extracted any layouts
        mock_autocad.get_single_layout_values.assert_not_called()


class TestPushToBoqWriteBack:
    def test_set_layouts_tables_id_called(self, push_util, mock_autocad, mock_odoo):
        mock_autocad.get_doc_layouts.return_value = ["L1"]
        mock_autocad.get_single_layout_values.return_value = {"layout_name": "L1", "detail": []}
        boq_result = [{"layout_name": "L1", "header_id": "H1"}]
        mock_odoo.import2boq.return_value = boq_result

        push_util.push_to_boq()

        mock_autocad.set_layouts_tables_id.assert_called_once_with(boq_result)
