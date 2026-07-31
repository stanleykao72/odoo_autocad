# -*- coding: utf-8 -*-
"""
Dispatcher 的 COM 跨線程路由測試。

COM 物件是 STA 封送的：在 GUI 主線程建立、從背景線程呼叫會得到
RPC_E_WRONG_THREAD (0x8001010E)，而 UtilAutoCAD 內部的 try/except 會把它
吞成空結果 —— 表現成「圖面沒有資料」。MCP server 正是跑在背景線程。

規則：
  * COM 模式 + 非主線程 + 已註冊代理 -> 繞回 GUI 主線程執行
  * 主線程 -> 直接呼叫（繞代理會死鎖）
  * IPC 模式 -> 直接呼叫（File IPC 無線程限制）
  * 未註冊代理 -> 直接呼叫（例如無 GUI 的 stdio 模式）
"""
import threading
from unittest.mock import MagicMock, patch

import pytest

from utility.util_autocad_dispatcher import UtilAutoCADDispatcher


@pytest.fixture(autouse=True)
def mock_backends():
    import utility.util_autocad as ua
    with patch.object(ua, 'pythoncom', MagicMock()), \
            patch.object(ua, 'client', MagicMock()):
        yield


@pytest.fixture
def dispatcher(mock_odoo_util, mock_log_util):
    d = UtilAutoCADDispatcher(mock_odoo_util, mock_log_util, mode="com")
    d._active = MagicMock()
    d._active.get_doc_layouts.return_value = ["配置1"]
    return d


def _proxy(result=None, success=True, error="boom"):
    p = MagicMock()
    p.execute_in_gui.return_value = (
        {"success": True, "result": result} if success
        else {"success": False, "error": error})
    return p


class TestRoutingDecision:
    def test_main_thread_calls_backend_directly(self, dispatcher):
        proxy = _proxy(["不該用到"])
        dispatcher.set_gui_proxy(proxy, threading.get_ident())

        assert dispatcher.get_doc_layouts() == ["配置1"]
        proxy.execute_in_gui.assert_not_called()

    def test_background_thread_goes_through_proxy(self, dispatcher):
        proxy = _proxy(["配置1"])
        dispatcher.set_gui_proxy(proxy, threading.get_ident() + 1)   # 非本線程

        assert dispatcher.get_doc_layouts() == ["配置1"]
        proxy.execute_in_gui.assert_called_once()
        kwargs = proxy.execute_in_gui.call_args.kwargs
        assert kwargs["method"] == "get_doc_layouts"

    def test_ipc_mode_never_uses_proxy(self, mock_odoo_util, mock_log_util):
        with patch("utility.util_autocad_ipc.UtilAutoCADIPC") as ipc_cls:
            d = UtilAutoCADDispatcher(mock_odoo_util, mock_log_util, mode="ipc")
            d._active = MagicMock()
            d._active.get_doc_layouts.return_value = ["L1"]
            proxy = _proxy(["不該用到"])
            d.set_gui_proxy(proxy, threading.get_ident() + 1)

            assert d.get_doc_layouts() == ["L1"]
            proxy.execute_in_gui.assert_not_called()

    def test_no_proxy_registered_calls_directly(self, dispatcher):
        """無 GUI 的 stdio 模式：沒有代理就直接呼叫"""
        assert dispatcher.get_doc_layouts() == ["配置1"]


class TestProxyPayload:
    def test_arguments_are_forwarded(self, dispatcher):
        proxy = _proxy(True)
        dispatcher.set_gui_proxy(proxy, threading.get_ident() + 1)

        dispatcher.set_block_attributes({"spec": "X"}, "配置1")

        kwargs = proxy.execute_in_gui.call_args.kwargs
        assert kwargs["method"] == "set_block_attributes"
        assert kwargs["args"] == ({"spec": "X"}, "配置1")

    def test_dict_result_is_unwrapped_correctly(self, dispatcher):
        """後端回傳 dict 時不可與代理協定混淆"""
        proxy = _proxy({"all": [{"layout_name": "01"}]})
        dispatcher.set_gui_proxy(proxy, threading.get_ident() + 1)

        assert dispatcher.get_layouts_values() == {"all": [{"layout_name": "01"}]}

    def test_generous_timeout_is_passed(self, dispatcher):
        """BOQ 擷取可能數十秒，不可沿用預設 10 秒"""
        proxy = _proxy({})
        dispatcher.set_gui_proxy(proxy, threading.get_ident() + 1)

        dispatcher.get_layouts_values()
        assert proxy.execute_in_gui.call_args.kwargs["timeout"] >= 60


class TestProxyFailure:
    def test_failure_raises_instead_of_returning_empty(self, dispatcher):
        """代理失敗必須拋錯 —— 回空值就會重演『看起來像沒資料』的問題"""
        proxy = _proxy(success=False, error="GUI代理執行超時")
        dispatcher.set_gui_proxy(proxy, threading.get_ident() + 1)

        with pytest.raises(RuntimeError, match="get_doc_layouts"):
            dispatcher.get_doc_layouts()


class TestCallBackendHandler:
    def test_handler_invokes_active_backend(self, dispatcher):
        from utility.util_gui_proxy import setup_backend_proxy, get_gui_proxy
        setup_backend_proxy(dispatcher, None)
        handler = get_gui_proxy().handlers["call_backend"]

        out = handler(method="get_doc_layouts", args=(), kwargs={})
        assert out == {"success": True, "result": ["配置1"]}

    def test_handler_reports_unknown_method(self, dispatcher):
        from utility.util_gui_proxy import setup_backend_proxy, get_gui_proxy
        dispatcher._active = MagicMock(spec=[])      # 沒有任何方法
        setup_backend_proxy(dispatcher, None)
        handler = get_gui_proxy().handlers["call_backend"]

        out = handler(method="does_not_exist", args=(), kwargs={})
        assert out["success"] is False
