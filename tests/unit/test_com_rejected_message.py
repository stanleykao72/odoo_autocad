# -*- coding: utf-8 -*-
"""
COM 連線被 AutoCAD 持續拒絕時的訊息測試。

背景：AutoCAD 以 RPC_E_CALL_REJECTED 拒絕所有 COM 呼叫時，AutoCAD 其實是
開著的、也有文件，舊訊息「請確認 AutoCAD 已開啟且有文件」會誤導使用者。
（已實測：此時 VBScript 等獨立 COM 客戶端同樣被拒，另開乾淨實例則正常。）
"""
from unittest.mock import MagicMock, patch

import pytest

RPC_E_CALL_REJECTED = -2147418111


@pytest.fixture(autouse=True)
def mock_com_handles():
    import utility.util_autocad as ua
    with patch.object(ua, 'pythoncom', MagicMock()) as mock_pythoncom, \
            patch.object(ua, 'client', MagicMock()), \
            patch.object(ua, 'time', MagicMock()):   # 跳過 retry 的 sleep
        mock_pythoncom.com_error = type('com_error', (Exception,), {})
        yield mock_pythoncom


@pytest.fixture
def backend(mock_odoo_util, mock_log_util):
    from utility.util_autocad import UtilAutoCAD
    return UtilAutoCAD(mock_odoo_util, mock_log_util)


def _logged(backend):
    return " ".join(c.args[0] for c in backend.log.safe_log_insert.call_args_list)


def _make_rejecting_app(mock_pythoncom, hresult):
    """GetActiveObject 成功，但之後每次呼叫都丟出指定 HRESULT"""
    err = mock_pythoncom.com_error(hresult, "rejected", None, None)
    err.args = (hresult, "rejected", None, None)
    raw = MagicMock()
    raw.QueryInterface.return_value.GetIDsOfNames.side_effect = err
    mock_pythoncom.GetActiveObject.return_value = raw
    return raw


class TestPersistentRejection:
    def test_message_does_not_blame_missing_document(self, backend, mock_com_handles):
        _make_rejecting_app(mock_com_handles, RPC_E_CALL_REJECTED)
        backend.connect_autocad(MagicMock())

        out = _logged(backend)
        assert "RPC_E_CALL_REJECTED" in out, "應指出實際的 COM 錯誤"
        assert "請確認 AutoCAD 已開啟且有文件" not in out, \
            "AutoCAD 開著且有文件，此訊息會誤導使用者"

    def test_message_offers_actionable_steps(self, backend, mock_com_handles):
        _make_rejecting_app(mock_com_handles, RPC_E_CALL_REJECTED)
        backend.connect_autocad(MagicMock())

        out = _logged(backend)
        assert "重新啟動 AutoCAD" in out
        assert "IPC" in out, "應提示可改用 IPC 模式"

    def test_not_connected_after_rejection(self, backend, mock_com_handles):
        _make_rejecting_app(mock_com_handles, RPC_E_CALL_REJECTED)
        backend.connect_autocad(MagicMock())
        assert backend.connected_autocad() is False
        assert backend.acad is None and backend.doc is None


class TestOtherComErrors:
    def test_non_rejection_error_keeps_generic_message(self, backend, mock_com_handles):
        """非 RPC_E_CALL_REJECTED 的錯誤不應套用「拒絕」那套說明"""
        _make_rejecting_app(mock_com_handles, -2147024891)  # E_ACCESSDENIED
        backend.connect_autocad(MagicMock())

        out = _logged(backend)
        assert "無法獲取 ActiveDocument" in out
        assert "重新啟動 AutoCAD" not in out
