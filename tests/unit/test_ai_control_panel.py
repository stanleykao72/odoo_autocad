# -*- coding: utf-8 -*-
"""
Unit tests for the MCP control panel in ModernFormMain.

註：舊版本測的是 TCP Socket + Named Pipe 那套 MCPServerManager
（toggle_mcp_server / update_mcp_status_display / start_all_servers），
該架構已由 utility/util_mcp_manager.py 的 MCPManager 取代，
UI 也改為 toggle_sse_server + on_mcp_status_update。
本檔已改寫為對照現行實作。
"""
from unittest.mock import Mock, patch

import pytest

from forms.form_main_modern import ModernFormMain


class TestMCPControlPanelMethods:
    """介面方法存在性"""

    @pytest.mark.parametrize("name", [
        "create_ai_control_banner",
        "create_sse_status_indicator",
        "create_sse_control_panel",
        "initialize_mcp_server_manager",
        "toggle_sse_server",
        "auto_start_mcp_server",
        "on_mcp_status_update",
        "show_sse_status",
        "test_sse_connection",
    ])
    def test_method_exists_and_callable(self, name):
        assert hasattr(ModernFormMain, name), f"{name} 應存在"
        assert callable(getattr(ModernFormMain, name)), f"{name} 應可呼叫"

    def test_removed_legacy_methods_are_gone(self):
        """舊的 TCP/Pipe 控制方法已隨架構移除，不應復活"""
        for name in ("toggle_mcp_server", "update_mcp_status_display"):
            assert not hasattr(ModernFormMain, name), \
                f"{name} 屬於已移除的 TCP/Pipe 架構"

    def test_backward_compatible_alias_kept(self):
        """auto_start_sse_server 是 auto_start_mcp_server 的相容別名"""
        assert ModernFormMain.auto_start_sse_server is ModernFormMain.auto_start_mcp_server


def _form_stub():
    """建立不啟動 Tk 的 ModernFormMain 替身"""
    form = Mock(spec=ModernFormMain)
    form.mcp_manager = Mock()
    form.log_util = Mock()
    form.autocad_mode = "com"
    return form


class TestToggleSSEServer:
    def test_starts_when_stopped(self):
        form = _form_stub()
        form.mcp_manager.is_running = False
        form.toggle_sse_server = ModernFormMain.toggle_sse_server.__get__(form)

        form.toggle_sse_server()

        form.mcp_manager.start_server.assert_called_once()
        form.mcp_manager.stop_server.assert_not_called()

    def test_stops_when_running(self):
        form = _form_stub()
        form.mcp_manager.is_running = True
        form.toggle_sse_server = ModernFormMain.toggle_sse_server.__get__(form)

        form.toggle_sse_server()

        form.mcp_manager.stop_server.assert_called_once()
        form.mcp_manager.start_server.assert_not_called()

    def test_error_is_reported_not_raised(self):
        form = _form_stub()
        form.mcp_manager.is_running = False
        form.mcp_manager.start_server.side_effect = OSError("port in use")
        form.toggle_sse_server = ModernFormMain.toggle_sse_server.__get__(form)

        with patch("forms.form_main_modern.messagebox") as mb:
            form.toggle_sse_server()  # 不可往外拋
            mb.showerror.assert_called_once()
        form.log_util.safe_log_insert.assert_called()


class TestAutoStartMCPServer:
    def test_calls_start_server(self):
        form = _form_stub()
        form.mcp_manager.port = 8084
        form.auto_start_mcp_server = ModernFormMain.auto_start_mcp_server.__get__(form)

        form.auto_start_mcp_server()

        form.mcp_manager.start_server.assert_called_once()

    def test_failure_is_swallowed_and_logged(self):
        form = _form_stub()
        form.mcp_manager.port = 8084
        form.mcp_manager.start_server.side_effect = RuntimeError("boom")
        form.auto_start_mcp_server = ModernFormMain.auto_start_mcp_server.__get__(form)

        form.auto_start_mcp_server()  # 啟動失敗不可讓 GUI 無法開啟

        logged = " ".join(str(c) for c in form.log_util.safe_log_insert.call_args_list)
        assert "Auto-start failed" in logged


class TestMCPStatusCallback:
    def test_status_update_is_marshalled_to_ui_thread(self):
        """on_mcp_status_update 由背景 thread 呼叫，必須透過 after() 回主線程"""
        form = _form_stub()
        form.on_mcp_status_update = ModernFormMain.on_mcp_status_update.__get__(form)

        form.on_mcp_status_update(True, "running")

        form.after.assert_called_once()
        assert form.after.call_args[0][0] == 0

    def test_running_state_updates_indicators(self):
        form = _form_stub()
        form.mcp_manager.port = 8084
        form.sse_status_label = Mock()
        form.sse_toggle_button = Mock()
        form.sse_info_label = Mock()
        form._log_mcp_connection_guide = Mock()
        # after(0, fn) → 直接執行 fn，模擬主線程
        form.after = Mock(side_effect=lambda delay, fn: fn())
        form.on_mcp_status_update = ModernFormMain.on_mcp_status_update.__get__(form)

        form.on_mcp_status_update(True, "running")

        form.sse_status_label.configure.assert_called_with(text="🟢")
        form.sse_toggle_button.configure.assert_called_with(text="⏹️")
        form.sse_info_label.configure.assert_called_with(text="MCP: :8084")
        form._log_mcp_connection_guide.assert_called_once()

    def test_stopped_state_updates_indicators(self):
        form = _form_stub()
        form.sse_status_label = Mock()
        form.sse_toggle_button = Mock()
        form.sse_info_label = Mock()
        form._log_mcp_connection_guide = Mock()
        form.after = Mock(side_effect=lambda delay, fn: fn())
        form.on_mcp_status_update = ModernFormMain.on_mcp_status_update.__get__(form)

        form.on_mcp_status_update(False, "stopped")

        form.sse_status_label.configure.assert_called_with(text="🔴")
        form.sse_toggle_button.configure.assert_called_with(text="🌊")
        form.sse_info_label.configure.assert_called_with(text="")
        form._log_mcp_connection_guide.assert_not_called()

    def test_legacy_callback_signature_still_works(self):
        """on_mcp_sse_status_update(message, is_running) 為舊簽章的相容包裝"""
        form = _form_stub()
        form.on_mcp_status_update = Mock()
        form.on_mcp_sse_status_update = ModernFormMain.on_mcp_sse_status_update.__get__(form)

        form.on_mcp_sse_status_update("hello", True)

        form.on_mcp_status_update.assert_called_once_with(True, "hello")
