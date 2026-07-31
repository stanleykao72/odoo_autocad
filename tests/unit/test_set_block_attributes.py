# -*- coding: utf-8 -*-
"""
set_block_attributes (COM) 的行為測試。

實際案例：圖面的標題欄屬性區塊放在 Model 空間，配置內只有視埠
（實測 配置1／配置2 各只有 2 個 AcDbViewport，沒有任何其他實體）。
舊版在這種情況下「什麼都沒寫卻回傳 True」，表單於是顯示更新成功，
圖面卻毫無變化。
（對照組：IPC 端的 AutoLISP 會回報 "No attribute block found in layout"。）

屬性必須寫進使用者指定的那個配置，不可退回 Model 空間 —— 寫到 Model 會讓
所有配置共用同一份值。配置內沒有屬性區塊時應據實報錯。
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


def _viewport():
    v = MagicMock()
    v.ObjectName = "AcDbViewport"
    return v


def _layout(name, entities):
    lay = MagicMock()
    lay.Name = name
    lay.Block = entities
    return lay


def _wire(backend, layouts):
    """layouts: dict of name -> entity list"""
    objs = {n: _layout(n, e) for n, e in layouts.items()}
    doc = backend.acad.ActiveDocument
    doc.Layouts.Item.side_effect = lambda n: objs[n]
    doc.Layouts.__iter__ = lambda self: iter(objs.values())
    return objs


def _title_attrs():
    return [
        _attr("project_name", "舊專案"),
        _attr("product_name", "SUS"),
        _attr("spec", "#304"),
    ]


class TestLayoutHasNoAttributeBlock:
    """配置內沒有屬性區塊時，必須報錯而非寫到別處"""

    def test_raises_and_does_not_touch_model(self, backend):
        """迴歸測試：舊版此時回傳 True，表單誤判為成功"""
        model_attrs = _title_attrs()
        model_block = _attr_block("A$C7a03b403", model_attrs)
        # 配置內只有視埠 —— 與實際圖面相同
        _wire(backend, {"Model": [model_block], "配置1": [_viewport(), _viewport()]})

        with pytest.raises(RuntimeError, match="找不到屬性區塊"):
            backend.set_block_attributes({"spec": "1100-H14"}, "配置1")

        assert model_attrs[2].TextString == "#304", \
            "不可退回寫入 Model —— 那會讓所有配置共用同一份值"

    def test_error_explains_the_cause(self, backend):
        _wire(backend, {"Model": [_attr_block("TB", _title_attrs())],
                        "配置1": [_viewport()]})
        with pytest.raises(RuntimeError) as exc:
            backend.set_block_attributes({"spec": "X"}, "配置1")
        msg = str(exc.value)
        assert "配置1" in msg
        assert "Model" in msg, "應說明標題欄放在 Model 時無法寫入的原因"


class TestLayoutHasAttributeBlock:
    """配置本身有屬性區塊時，正常寫入且不動到 Model"""

    def test_writes_to_layout_block(self, backend):
        layout_attrs = _title_attrs()
        model_attrs = _title_attrs()
        _wire(backend, {"Model": [_attr_block("MODEL_TB", model_attrs)],
                        "配置1": [_viewport(), _attr_block("LAY_TB", layout_attrs)]})

        assert backend.set_block_attributes({"spec": "1100-H14"}, "配置1") is True
        assert layout_attrs[2].TextString == "1100-H14"
        assert model_attrs[2].TextString == "#304", "不可波及 Model 的區塊"

    def test_writes_multiple_fields(self, backend):
        layout_attrs = _title_attrs()
        _wire(backend, {"Model": [], "配置1": [_attr_block("TB", layout_attrs)]})

        backend.set_block_attributes(
            {"product_name": "1F大門更換地鉸鏈(人工)", "spec": "1100-H14"}, "配置1")
        assert layout_attrs[1].TextString == "1F大門更換地鉸鏈(人工)"
        assert layout_attrs[2].TextString == "1100-H14"

    def test_raises_when_no_requested_tag_matches(self, backend):
        """區塊存在但沒有任何一個要寫的欄位 —— 同樣不可靜默成功"""
        _wire(backend, {"Model": [], "配置1": [_attr_block("TB", _title_attrs())]})
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
