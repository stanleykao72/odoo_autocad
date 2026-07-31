# -*- coding: utf-8 -*-
"""
Tests for Odoo connection error reporting (utility/util_odoo.py).

重點：連線失敗時日誌必須明確指出原因，特別是 token 不正確的情況 ——
舊版一律記成「無效的 Swagger 文件」，401/403 也是同一句，使用者無從判斷。
另外 url 內含 API token，不可原樣寫進日誌。
"""
from unittest.mock import Mock, patch

import pytest
import requests
import simplejson

from utility.util_odoo import UtilOdoo

URL = "https://demo.odoo.com/api/v1/swagger.json?token=abcdef0123456789&db=demo"
CONN = {"host": "demo.odoo.com", "db_name": "demo", "url": URL, "token": "tok-1234567890"}


@pytest.fixture
def log():
    return Mock()


def _logged(log):
    """所有日誌內容（含連線參數行）"""
    return " ".join(c.args[0] for c in log.safe_log_insert.call_args_list)


def _connect_raising(exc, log):
    """建立 UtilOdoo 並讓 SwaggerClient.from_url 拋出 exc。

    回傳「錯誤訊息」行（以 ❌ 開頭）而非全部日誌 —— 連線參數行本來就會
    出現 "url:" / "token:" 等字樣，拿整份日誌做斷言會誤判。
    """
    with patch("utility.util_odoo.SwaggerClient.from_url", side_effect=exc):
        with pytest.raises(Exception):
            UtilOdoo(CONN, log)
    errors = [c.args[0] for c in log.safe_log_insert.call_args_list
              if c.args and c.args[0].startswith("❌")]
    assert errors, f"連線失敗必須留下 ❌ 錯誤訊息，實際日誌: {_logged(log)}"
    return " ".join(errors)


def _http_error(status):
    """模擬帶有 HTTP 狀態碼的錯誤（bravado 形態：例外自身帶 status_code）"""
    err = requests.exceptions.HTTPError(f"HTTP {status}")
    err.response = Mock(status_code=status)
    return err


class TestTokenErrorsAreNamed:
    @pytest.mark.parametrize("status", [401, 403])
    def test_auth_failure_mentions_token(self, status, log):
        out = _connect_raising(_http_error(status), log)
        assert "token" in out.lower(), f"HTTP {status} 應明確指出 token 問題，實際: {out}"
        assert str(status) in out

    def test_401_tells_user_where_to_fix(self, log):
        out = _connect_raising(_http_error(401), log)
        assert "token.yaml" in out

    def test_json_decode_error_suggests_token(self, log):
        """token 錯誤時 Odoo 常導向登入頁 → JSON 解析失敗，而非回 401"""
        exc = simplejson.errors.JSONDecodeError("Expecting value", "<html>", 0)
        out = _connect_raising(exc, log)
        assert "token" in out.lower()


class TestOtherFailuresAreDistinguished:
    def test_connection_error(self, log):
        out = _connect_raising(requests.exceptions.ConnectionError(), log)
        assert "無法連線" in out
        assert "token" not in out.lower(), "網路問題不可誤報成 token 問題"

    def test_timeout(self, log):
        out = _connect_raising(requests.exceptions.Timeout(), log)
        assert "逾時" in out
        assert "token" not in out.lower()

    def test_404_is_not_reported_as_token_problem(self, log):
        out = _connect_raising(_http_error(404), log)
        assert "404" in out
        assert "token" not in out.lower()

    def test_500_is_not_reported_as_token_problem(self, log):
        out = _connect_raising(_http_error(500), log)
        assert "500" in out
        assert "token" not in out.lower()

    def test_unknown_status_still_reported(self, log):
        out = _connect_raising(_http_error(418), log)
        assert "418" in out


class TestSecretsNotLogged:
    """這裡刻意檢查「整份日誌」，包含連線參數那幾行"""

    def test_url_token_is_masked(self, log):
        _connect_raising(requests.exceptions.ConnectionError(), log)
        out = _logged(log)
        assert "abcdef0123456789" not in out, "url 中的 API token 不可寫入日誌"
        assert "demo.odoo.com" in out, "非機密部分仍應可見，以利除錯"

    def test_user_token_is_masked(self, log):
        _connect_raising(requests.exceptions.ConnectionError(), log)
        assert "tok-1234567890" not in _logged(log)


class TestDisconnectedConstruction:
    def test_connect_false_does_not_call_network(self, log):
        with patch("utility.util_odoo.SwaggerClient.from_url") as m:
            util = UtilOdoo(CONN, log, connect=False)
        m.assert_not_called()
        assert util.connected_odoo() is False
        assert util.odoo is None

    def test_disconnected_instance_can_retry_later(self, log):
        util = UtilOdoo(CONN, log, connect=False)
        with patch("utility.util_odoo.SwaggerClient.from_url", return_value=Mock()):
            util.connect_odoo(CONN)
        assert util.connected_odoo() is True

    def test_connect_true_is_the_default_and_raises(self, log):
        """預設仍 fail fast，避免其他呼叫端拿到半殘實例卻不自知"""
        with patch("utility.util_odoo.SwaggerClient.from_url",
                   side_effect=requests.exceptions.ConnectionError()):
            with pytest.raises(requests.exceptions.ConnectionError):
                UtilOdoo(CONN, log)
