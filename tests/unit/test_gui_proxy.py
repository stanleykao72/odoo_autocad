# -*- coding: utf-8 -*-
"""
Tests for GUIProxy (utility/util_gui_proxy.py).

Verifies:
- 請求／回覆往返正確
- 未註冊的 action 回傳錯誤而非拋例外
- handler 拋例外時被包成錯誤回應
- 超時後不會殘留未取走的響應（記憶體洩漏迴歸測試）
"""
import queue
import threading

import pytest

from utility.util_gui_proxy import GUIProxy


@pytest.fixture
def proxy():
    return GUIProxy(timeout=1.0)


def _pump(proxy, times=1, delay=0.0):
    """在背景模擬 GUI 主線程，呼叫 process_requests()"""
    def _run():
        import time
        for _ in range(times):
            if delay:
                time.sleep(delay)
            proxy.process_requests()
    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return t


class TestGUIProxyRoundTrip:
    def test_handler_result_returned_to_caller(self, proxy):
        proxy.register_handler("ping", lambda: {"success": True, "pong": 1})
        _pump(proxy, times=20, delay=0.01)

        result = proxy.execute_in_gui("ping")
        assert result["success"] is True
        assert result["pong"] == 1

    def test_params_are_forwarded(self, proxy):
        proxy.register_handler("echo", lambda value: {"success": True, "value": value})
        _pump(proxy, times=20, delay=0.01)

        result = proxy.execute_in_gui("echo", value="hello")
        assert result["value"] == "hello"

    def test_non_dict_result_is_wrapped(self, proxy):
        proxy.register_handler("num", lambda: 42)
        _pump(proxy, times=20, delay=0.01)

        result = proxy.execute_in_gui("num")
        assert result == {"success": True, "result": 42}


class TestGUIProxyErrors:
    def test_unknown_action(self, proxy):
        proxy.register_handler("known", lambda: {"success": True})
        _pump(proxy, times=20, delay=0.01)

        result = proxy.execute_in_gui("unknown")
        assert result["success"] is False
        assert "unknown" in result["error"]
        assert "known" in result["available_actions"]

    def test_handler_exception_is_captured(self, proxy):
        def boom():
            raise ValueError("炸了")
        proxy.register_handler("boom", boom)
        _pump(proxy, times=20, delay=0.01)

        result = proxy.execute_in_gui("boom")
        assert result["success"] is False
        assert result["exception_type"] == "ValueError"


class TestGUIProxyTimeout:
    def test_timeout_returns_error(self, proxy):
        proxy.register_handler("slow", lambda: {"success": True})
        # 不啟動 pump — 沒有 GUI 線程處理請求
        result = proxy.execute_in_gui("slow", timeout=0.2)
        assert result["success"] is False
        assert result["timeout"] is True

    def test_late_response_does_not_leak(self, proxy):
        """迴歸測試：呼叫端超時離開後，遲到的響應必須被丟棄而非無限累積。

        舊版把響應存進共用的 self.response_cache，超時後沒有人 pop，
        長時間執行會持續累積記憶體。
        """
        proxy.register_handler("slow", lambda: {"success": True})

        result = proxy.execute_in_gui("slow", timeout=0.1)
        assert result["timeout"] is True

        # GUI 線程「稍後」才處理這個請求 — 不可拋例外
        processed = proxy.process_requests()
        assert processed == 1

        # 沒有任何共用的響應暫存區殘留
        assert not hasattr(proxy, "response_cache")
        assert proxy.request_queue.empty()


class TestGUIProxyQueueSemantics:
    def test_each_request_has_own_reply_queue(self, proxy):
        proxy.register_handler("a", lambda: {"success": True, "who": "a"})
        proxy.register_handler("b", lambda: {"success": True, "who": "b"})
        _pump(proxy, times=40, delay=0.01)

        results = {}

        def call(action):
            results[action] = proxy.execute_in_gui(action)

        threads = [threading.Thread(target=call, args=(a,)) for a in ("a", "b")]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=5)

        # 回覆不可互相串線
        assert results["a"]["who"] == "a"
        assert results["b"]["who"] == "b"
