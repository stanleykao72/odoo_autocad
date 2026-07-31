# -*- coding: utf-8 -*-
"""
set_block_attributes (COM) 的行為測試。

實際案例：圖面的標題欄屬性區塊放在 Model 空間，配置只是透過視埠看到它，
因此配置本身沒有任何含屬性的區塊。舊版在這種情況下「什麼都沒寫卻回傳 True」，
表單於是顯示更新成功，圖面卻毫無變化。
（對照組：IPC 端的 AutoLISP 會回報 "No attribute block found in layout"。）
"""
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture(autouse=True)
def mock_com_handles():
    import utility.util_autocad as ua
    with patch.object(ua, 'pythoncom', MagicMock()), \
            patch.object(ua, 'client', MagicMock()):
        yield


@pytest.fixture
def backend(mock_odoo_util, mock_log_util):
    from utility.util_autocad import UtilAutoCAD
    util = UtilAutoCAD(mock_odoo_util, mock_log_util)
    util.acad = MagicMock()
    return util


def _attr(tag, value):
    a = MagicMock()
    a.TagString = tag
    a.TextString = value
    return a


def _attr_block(name, attrs):
    b = MagicMock()
    b.ObjectName = "AcDbBlockReference"
    b.Name = name
    b.HasAttributes = True
    b.GetAttributes.return_value = attrs
    return b


def _layout(name, blocks):
    lay = MagicMock()
    lay.Name = name
    lay.Block = blocks
    return lay


def _wire(backend, layouts):
    """layouts: dict of name -> block list"""
    objs = {n: _layout(n, b) for n, b in layouts.items()}
    doc = backend.acad.ActiveDocument
    doc.Layouts.Item.side_effect = lambda n: objs[n]
    doc.Layouts.__iter__ = lambda self: iter(objs.values())
    return objs


TITLE_ATTRS = lambda: [                      # noqa: E731
    _attr("project_name", "舊專案"),
    _attr("product_name", "SUS"),
    _attr("spec", "#304"),
]


class TestLayoutHasNoAttributeBlock:
    """配置內沒有屬性區塊時的行為"""

    def test_falls_back_to_model_block(self, backend):
        model_block = _attr_block("A$C7a03b403", TITLE_ATTRS())
        _wire(backend, {"Model": [model_block], "配置1": []})

        assert backend.set_block_attributes({"spec": "1100-H14"}, "配置1") is True
        assert model_block.GetAttributes.return_value[2].TextString == "1100-H14"

    def test_fallback_is_logged(self, backend):
        _wire(backend, {"Model": [_attr_block("TB", TITLE_ATTRS())], "配置1": []})
        backend.set_block_attributes({"spec": "X"}, "配置1")

        logged = " ".join(c.args[0] for c in backend.log.safe_log_insert.call_args_list)
        assert "Model" in logged, "退回 Model 寫入必須留下紀錄"

    def test_raises_when_neither_layout_nor_model_has_block(self, backend):
        """迴歸測試：舊版此時回傳 True，表單誤判為成功"""
        _wire(backend, {"Model": [], "配置1": []})
        with pytest.raises(RuntimeError, match="找不到屬性區塊"):
            backend.set_block_attributes({"spec": "X"}, "配置1")


class TestLayoutHasAttributeBlock:
    """配置本身有屬性區塊時，不應動到 Model"""

    def test_writes_to_layout_block(self, backend):
        layout_block = _attr_block("LAY_TB", TITLE_ATTRS())
        model_block = _attr_block("MODEL_TB", TITLE_ATTRS())
        _wire(backend, {"Model": [model_block], "配置1": [layout_block]})

        assert backend.set_block_attributes({"spec": "1100-H14"}, "配置1") is True
        assert layout_block.GetAttributes.return_value[2].TextString == "1100-H14"
        assert model_block.GetAttributes.return_value[2].TextString == "#304", \
            "配置有自己的屬性區塊時不可寫到 Model"

    def test_raises_when_no_requested_tag_matches(self, backend):
        """區塊存在但沒有任何一個要寫的欄位 —— 同樣不可靜默成功"""
        _wire(backend, {"Model": [], "配置1": [_attr_block("TB", TITLE_ATTRS())]})
        with pytest.raises(RuntimeError, match="沒有對應的欄位"):
            backend.set_block_attributes({"不存在的欄位": "X"}, "配置1")


class TestPollingDoesNotSpamLog:
    def test_get_active_layout_is_silent(self, backend):
        """COM 模式每 3 秒輪詢一次，此方法不可每次都寫日誌"""
        layout = MagicMock()
        layout.Name = "配置1"
        backend.doc = MagicMock()
        backend.doc.ActiveLayout = layout

        backend.log.safe_log_insert.reset_mock()
        for _ in range(5):
            assert backend.get_active_layout() == "配置1"
        assert backend.log.safe_log_insert.call_count == 0, \
            "輪詢呼叫不可寫入日誌，否則會洗掉真正的訊息"
