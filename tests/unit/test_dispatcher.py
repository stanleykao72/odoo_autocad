# -*- coding: utf-8 -*-
"""
Tests for UtilAutoCADDispatcher.

Verifies:
- Routes calls to active backend
- get_single_layout_values() no longer uses hasattr (direct delegation)
- Mode switching works
- _process_pr_no_ipc only runs in IPC mode
"""
import pytest
import sys
from unittest.mock import MagicMock, patch


@pytest.fixture(autouse=True)
def mock_win32com_modules():
    mods = {}
    for name in ('win32com', 'win32com.client', 'pythoncom'):
        if name not in sys.modules:
            mods[name] = MagicMock()
            sys.modules[name] = mods[name]
    yield
    for name, mod in mods.items():
        if sys.modules.get(name) is mod:
            del sys.modules[name]


@pytest.fixture
def dispatcher(mock_odoo_util, mock_log_util):
    from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
    d = UtilAutoCADDispatcher(mock_odoo_util, mock_log_util, mode="com")
    return d


class TestDispatcherRouting:
    def test_delegates_get_active_layout(self, dispatcher):
        dispatcher._active.get_active_layout = MagicMock(return_value="Layout1")
        result = dispatcher.get_active_layout()
        assert result == "Layout1"
        dispatcher._active.get_active_layout.assert_called_once()

    def test_delegates_get_doc_layouts(self, dispatcher):
        dispatcher._active.get_doc_layouts = MagicMock(return_value=["L1", "L2"])
        result = dispatcher.get_doc_layouts()
        assert result == ["L1", "L2"]

    def test_delegates_get_layouts_values(self, dispatcher):
        dispatcher._active.get_layouts_values = MagicMock(return_value={"all": []})
        result = dispatcher.get_layouts_values()
        assert result == {"all": []}

    def test_delegates_get_single_layout_values_directly(self, dispatcher):
        """get_single_layout_values no longer uses hasattr — direct call"""
        dispatcher._active.get_single_layout_values = MagicMock(return_value={"layout_name": "L1"})
        result = dispatcher.get_single_layout_values("L1")
        assert result == {"layout_name": "L1"}
        dispatcher._active.get_single_layout_values.assert_called_once_with("L1")

    def test_delegates_get_block_attributes(self, dispatcher):
        dispatcher._active.get_block_attributes = MagicMock(return_value={"pr_no": "PR001"})
        result = dispatcher.get_block_attributes()
        assert result == {"pr_no": "PR001"}

    def test_delegates_set_block_attributes(self, dispatcher):
        dispatcher._active.set_block_attributes = MagicMock(return_value=True)
        result = dispatcher.set_block_attributes({"pr_no": "X"}, "Layout1")
        assert result is True

    def test_delegates_get_layouts_header_id_to_pr(self, dispatcher):
        dispatcher._active.get_layouts_header_id_to_pr = MagicMock(return_value={"all": ["H1"]})
        result = dispatcher.get_layouts_header_id_to_pr()
        assert result == {"all": ["H1"]}


class TestDispatcherNoHasattr:
    def test_no_hasattr_in_get_single_layout_values(self):
        """Verify source code doesn't use hasattr for get_single_layout_values"""
        import inspect
        from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
        source = inspect.getsource(UtilAutoCADDispatcher.get_single_layout_values)
        assert "hasattr" not in source


class TestDispatcherModeSwitch:
    def test_switch_com_to_ipc(self, dispatcher, mock_odoo_util, mock_log_util):
        assert dispatcher.mode == "com"
        # Patch at the module level where the import happens
        with patch.dict('sys.modules', {'utility.util_autocad_ipc': MagicMock()}):
            from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
            # Re-create with fresh import context
            d = UtilAutoCADDispatcher(mock_odoo_util, mock_log_util, mode="com")
            d.switch_mode("ipc")
            assert d.mode == "ipc"

    def test_switch_same_mode_is_noop(self, dispatcher):
        old_active = dispatcher._active
        dispatcher.switch_mode("com")
        assert dispatcher._active is old_active

    def test_ipc_connect_calls_process_pr_no(self, mock_odoo_util, mock_log_util):
        """connect_autocad in IPC mode calls _process_pr_no_ipc"""
        from utility.util_autocad_dispatcher import UtilAutoCADDispatcher

        # Create dispatcher with a mock IPC backend injected directly
        d = UtilAutoCADDispatcher.__new__(UtilAutoCADDispatcher)
        d.odoo_util = mock_odoo_util
        d.log_util = mock_log_util
        d._mode = "ipc"
        d._com_backend = None
        d._com_progid = "AutoCAD.Application"
        d._ipc_target_hwnd = None

        mock_ipc = MagicMock()
        mock_ipc.connected_autocad.return_value = True
        d._ipc_backend = mock_ipc
        d._active = mock_ipc
        d._process_pr_no_ipc = MagicMock()

        d.connect_autocad()

        d._process_pr_no_ipc.assert_called_once()


class TestDispatcherProperties:
    def test_project_id_setter(self, dispatcher):
        dispatcher.project_id = 42
        assert dispatcher._active.project_id == 42

    def test_pr_no_setter(self, dispatcher):
        dispatcher.pr_no = "PR001"
        assert dispatcher._active.pr_no == "PR001"
