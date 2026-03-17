# -*- coding: utf-8 -*-
"""
Tests for AutoCADBackendInterface ABC contract.

Verifies:
- Both backends inherit ABC
- Missing methods → TypeError on instantiation
- All 17 abstract methods exist on both backends
"""
import pytest
import sys
from unittest.mock import MagicMock, patch
import inspect


# Mock win32com before importing COM backend
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


class TestABCContract:
    """ABC interface contract tests"""

    def test_abc_has_17_abstract_methods(self):
        from utility.autocad_backend_interface import AutoCADBackendInterface
        abstract = {
            name for name, method in inspect.getmembers(AutoCADBackendInterface)
            if getattr(method, '__isabstractmethod__', False)
        }
        assert len(abstract) == 17
        expected = {
            'connected_autocad', 'connect_autocad',
            'get_active_layout', 'get_doc_layouts',
            'get_layouts_values', 'get_single_layout_values',
            'get_block_attributes', 'set_block_attributes',
            'set_layouts_tables_id', 'get_layouts_header_id_to_pr',
            'clear_table_id', 'clear_all_tables_id',
            'draw_line', 'draw_circle', 'set_layer', 'list_layers',
            'scan_elements',
        }
        assert abstract == expected

    def test_com_inherits_abc(self):
        from utility.autocad_backend_interface import AutoCADBackendInterface
        from utility.util_autocad import UtilAutoCAD
        assert issubclass(UtilAutoCAD, AutoCADBackendInterface)

    def test_ipc_inherits_abc(self):
        from utility.autocad_backend_interface import AutoCADBackendInterface
        from utility.util_autocad_ipc import UtilAutoCADIPC
        assert issubclass(UtilAutoCADIPC, AutoCADBackendInterface)

    def test_com_instantiable(self, mock_odoo_util, mock_log_util):
        """COM backend can be instantiated (all abstract methods implemented)"""
        from utility.util_autocad import UtilAutoCAD
        instance = UtilAutoCAD(mock_odoo_util, mock_log_util)
        assert instance is not None

    def test_ipc_instantiable(self, mock_log_util):
        """IPC backend can be instantiated (all abstract methods implemented)"""
        from utility.util_autocad_ipc import UtilAutoCADIPC
        instance = UtilAutoCADIPC(log_util=mock_log_util)
        assert instance is not None

    def test_incomplete_subclass_raises_typeerror(self):
        """A subclass missing abstract methods cannot be instantiated"""
        from utility.autocad_backend_interface import AutoCADBackendInterface

        class IncompleteBackend(AutoCADBackendInterface):
            pass  # no methods implemented

        with pytest.raises(TypeError):
            IncompleteBackend()

    def test_com_has_all_abc_methods(self, mock_odoo_util, mock_log_util):
        from utility.autocad_backend_interface import AutoCADBackendInterface
        from utility.util_autocad import UtilAutoCAD

        instance = UtilAutoCAD(mock_odoo_util, mock_log_util)
        abstract = {
            name for name, method in inspect.getmembers(AutoCADBackendInterface)
            if getattr(method, '__isabstractmethod__', False)
        }
        for method_name in abstract:
            assert hasattr(instance, method_name), f"COM missing: {method_name}"
            assert callable(getattr(instance, method_name))

    def test_ipc_has_all_abc_methods(self, mock_log_util):
        from utility.autocad_backend_interface import AutoCADBackendInterface
        from utility.util_autocad_ipc import UtilAutoCADIPC

        instance = UtilAutoCADIPC(log_util=mock_log_util)
        abstract = {
            name for name, method in inspect.getmembers(AutoCADBackendInterface)
            if getattr(method, '__isabstractmethod__', False)
        }
        for method_name in abstract:
            assert hasattr(instance, method_name), f"IPC missing: {method_name}"
            assert callable(getattr(instance, method_name))
