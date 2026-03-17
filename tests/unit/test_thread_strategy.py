# -*- coding: utf-8 -*-
"""
Tests for COM/IPC thread strategy in form_main_modern.py.

Verifies via source-code inspection (no Tk runtime required):
- COM mode: _run_with_progress_main_thread does NOT use threading.Thread
- IPC mode: _run_with_progress_background DOES use threading.Thread
- _run_with_progress dispatches by autocad_mode
"""
import pytest
import os


@pytest.fixture
def form_source():
    """Read form_main_modern.py source as text"""
    path = os.path.join(os.path.dirname(__file__), "..", "..", "forms", "form_main_modern.py")
    with open(path, encoding="utf-8") as f:
        return f.read()


def _extract_method(source, method_name):
    """Extract a method body from source text (simple indent-based)"""
    lines = source.split("\n")
    start = None
    for i, line in enumerate(lines):
        if f"def {method_name}(" in line:
            start = i
            break
    if start is None:
        return ""
    # Find the indentation of def line
    indent = len(lines[start]) - len(lines[start].lstrip())
    body_lines = [lines[start]]
    for line in lines[start + 1:]:
        stripped = line.lstrip()
        if stripped == "" or stripped.startswith("#"):
            body_lines.append(line)
            continue
        line_indent = len(line) - len(stripped)
        if line_indent <= indent and stripped:
            break
        body_lines.append(line)
    return "\n".join(body_lines)


class TestThreadStrategySourceCode:
    """Verify source code structure"""

    def test_com_main_thread_no_threading(self, form_source):
        """COM path must not use threading.Thread"""
        method = _extract_method(form_source, "_run_with_progress_main_thread")
        assert method, "_run_with_progress_main_thread not found"
        assert "threading.Thread" not in method
        assert "Thread(" not in method

    def test_ipc_background_uses_threading(self, form_source):
        """IPC path must use threading.Thread"""
        method = _extract_method(form_source, "_run_with_progress_background")
        assert method, "_run_with_progress_background not found"
        assert "threading.Thread" in method

    def test_run_with_progress_dispatches_by_mode(self, form_source):
        """_run_with_progress dispatches based on autocad_mode"""
        method = _extract_method(form_source, "_run_with_progress")
        assert "autocad_mode" in method
        assert "_run_with_progress_main_thread" in method
        assert "_run_with_progress_background" in method

    def test_com_mode_calls_main_thread(self, form_source):
        """COM mode routes to _run_with_progress_main_thread"""
        method = _extract_method(form_source, "_run_with_progress")
        assert '"com"' in method


class TestThreadStrategySeparation:
    """Verify the two methods exist"""

    def test_main_thread_method_exists(self, form_source):
        assert "def _run_with_progress_main_thread(" in form_source

    def test_background_method_exists(self, form_source):
        assert "def _run_with_progress_background(" in form_source

    def test_both_have_same_params(self, form_source):
        """Both methods accept (title, message, worker, success_msg, fail_msg)"""
        for method in ("_run_with_progress_main_thread", "_run_with_progress_background"):
            sig_line = _extract_method(form_source, method).split("\n")[0]
            for param in ("title", "message", "worker", "success_msg", "fail_msg"):
                assert param in sig_line, f"{method} missing param: {param}"

    def test_com_main_thread_uses_update_idletasks(self, form_source):
        """COM path uses update_idletasks() for UI responsiveness"""
        method = _extract_method(form_source, "_run_with_progress_main_thread")
        assert "update_idletasks" in method
