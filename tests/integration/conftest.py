# -*- coding: utf-8 -*-
"""
Shared fixtures for integration tests that require a live AutoCAD instance.

Usage:
    pytest tests/integration/ -m autocad      # COM tests (need Full AutoCAD)
    pytest tests/integration/ -m ipc          # IPC tests (need AutoCAD LT/Full + McpDispatch.vlx)
    pytest tests/integration/ -m "autocad or ipc"  # both
"""
import pytest
import sys
import os
import logging

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


class _SimpleLog:
    """Minimal logger compatible with UtilLog.safe_log_insert"""
    def __init__(self):
        self._logger = logging.getLogger("integration_test")
        self._logger.setLevel(logging.DEBUG)
        if not self._logger.handlers:
            h = logging.StreamHandler()
            h.setFormatter(logging.Formatter("%(message)s"))
            self._logger.addHandler(h)

    def safe_log_insert(self, msg):
        self._logger.info(msg.rstrip("\n"))

    # UtilLog compat
    info = safe_log_insert


@pytest.fixture(scope="session")
def log_util():
    return _SimpleLog()


# ── COM fixtures ─────────────────────────────────────────────

@pytest.fixture(scope="session")
def com_backend(log_util):
    """Live COM backend — requires Full AutoCAD running with an open .dwg"""
    try:
        import pythoncom
        from win32com import client
    except ImportError:
        pytest.skip("pywin32 not installed")

    from utility.util_autocad import UtilAutoCAD
    # Use a dummy odoo_util (Odoo not needed for pure AutoCAD tests)
    util = UtilAutoCAD(odoo_util=None, log_util=log_util)
    try:
        pythoncom.CoInitialize()
        util.acad = client.GetActiveObject("AutoCAD.Application")
        util.doc = util.acad.ActiveDocument
        if not util.doc:
            pytest.skip("AutoCAD has no open document")
    except Exception as e:
        pytest.skip(f"Cannot connect to AutoCAD COM: {e}")

    yield util


# ── IPC fixtures ─────────────────────────────────────────────

@pytest.fixture(scope="session")
def ipc_backend(log_util):
    """Live IPC backend — requires AutoCAD (LT or Full) with McpDispatch.vlx loaded"""
    from utility.util_autocad_ipc import UtilAutoCADIPC

    util = UtilAutoCADIPC(log_util=log_util)
    try:
        util.connect_autocad()
        if not util.connected_autocad():
            pytest.skip("AutoCAD IPC not responding (is McpDispatch.vlx loaded?)")
    except Exception as e:
        pytest.skip(f"Cannot connect to AutoCAD IPC: {e}")

    yield util


# ── Dispatcher fixtures ──────────────────────────────────────

@pytest.fixture(scope="session")
def com_dispatcher(log_util):
    """Dispatcher in COM mode"""
    from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
    try:
        d = UtilAutoCADDispatcher(odoo_util=None, log_util=log_util, mode="com")
        d.connect_autocad(main_body=None)
        if not d.connected_autocad():
            pytest.skip("COM dispatcher cannot connect")
    except Exception as e:
        pytest.skip(f"COM dispatcher init failed: {e}")
    yield d


@pytest.fixture(scope="session")
def ipc_dispatcher(log_util):
    """Dispatcher in IPC mode"""
    from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
    try:
        d = UtilAutoCADDispatcher(odoo_util=None, log_util=log_util, mode="ipc")
        d.connect_autocad()
        if not d.connected_autocad():
            pytest.skip("IPC dispatcher cannot connect")
    except Exception as e:
        pytest.skip(f"IPC dispatcher init failed: {e}")
    yield d
