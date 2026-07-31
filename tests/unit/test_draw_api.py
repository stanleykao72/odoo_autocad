# -*- coding: utf-8 -*-
"""
繪圖 API 測試（draw_line / draw_circle）。

兩個實測發現的問題：

1. 座標必須是 VARIANT(VT_ARRAY|VT_R8)
   舊版直接把 Python list 傳給 AddCircle/AddLine，AutoCAD 回
   0x80020009 (inner 0x80070057 = E_INVALIDARG)，也就是這些方法從來
   不能用。因為沒有任何呼叫端，問題一直沒浮現。

2. 目標空間不可寫死 ModelSpace
   本專案的圖面內容多半放在配置的圖紙空間（實測 2026.07.31-TEST.dwg
   的 Model 幾乎是空的），畫到 Model 等於看不見。預設改為作用中的配置。
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
    util.doc = MagicMock()

    active = MagicMock(); active.Name = "01"
    other = MagicMock(); other.Name = "配置2"
    util.doc.ActiveLayout = active
    util.doc.Layouts.Item.side_effect = lambda n: {"01": active, "配置2": other}[n]
    util._ensure_layer_exists = MagicMock()
    util._spaces = {"01": active.Block, "配置2": other.Block}
    return util


class TestCoordinateConversion:
    def test_to_point_uses_variant(self):
        """迴歸測試：舊版直接傳 list -> E_INVALIDARG"""
        import utility.util_autocad as ua
        with patch.object(ua, 'VARIANT') as variant:
            ua.UtilAutoCAD._to_point([1, 2, 3])
        variant.assert_called_once()
        args = variant.call_args.args
        assert args[1] == [1.0, 2.0, 3.0], "座標須轉為 float 陣列"

    def test_to_point_pads_to_three(self):
        import utility.util_autocad as ua
        with patch.object(ua, 'VARIANT') as variant:
            ua.UtilAutoCAD._to_point([5, 6])
        assert variant.call_args.args[1] == [5.0, 6.0, 0.0]

    def test_to_point_truncates_extra(self):
        import utility.util_autocad as ua
        with patch.object(ua, 'VARIANT') as variant:
            ua.UtilAutoCAD._to_point([1, 2, 3, 4])
        assert variant.call_args.args[1] == [1.0, 2.0, 3.0]

    def test_draw_circle_passes_variant_not_list(self, backend):
        backend.draw_circle([10, 20], 5)
        pt = backend.doc.ActiveLayout.Block.AddCircle.call_args.args[0]
        assert not isinstance(pt, list), "不可直接傳 list 給 COM"


class TestTargetSpace:
    def test_defaults_to_active_layout_not_model(self, backend):
        """迴歸測試：舊版寫死 ModelSpace，畫了看不到"""
        result = backend.draw_circle([0, 0], 20)

        backend.doc.ActiveLayout.Block.AddCircle.assert_called_once()
        backend.doc.ModelSpace.AddCircle.assert_not_called()
        assert result["layout_name"] == "01"

    def test_explicit_layout_is_used(self, backend):
        result = backend.draw_circle([0, 0], 20, layout_name="配置2")

        backend.doc.Layouts.Item.assert_called_with("配置2")
        backend._spaces["配置2"].AddCircle.assert_called_once()
        backend._spaces["01"].AddCircle.assert_not_called()
        assert result["layout_name"] == "配置2"

    def test_draw_line_also_uses_active_layout(self, backend):
        result = backend.draw_line([0, 0], [1, 1])

        backend.doc.ActiveLayout.Block.AddLine.assert_called_once()
        backend.doc.ModelSpace.AddLine.assert_not_called()
        assert result["layout_name"] == "01"


class TestSignatureCompatibility:
    def test_abc_declares_layout_name(self):
        import inspect
        from utility.autocad_backend_interface import AutoCADBackendInterface
        for name in ("draw_line", "draw_circle"):
            sig = inspect.signature(getattr(AutoCADBackendInterface, name))
            assert "layout_name" in sig.parameters, f"{name} 應接受 layout_name"

    def test_ipc_backend_accepts_layout_name(self):
        import inspect
        from utility.util_autocad_ipc import UtilAutoCADIPC
        for name in ("draw_line", "draw_circle"):
            sig = inspect.signature(getattr(UtilAutoCADIPC, name))
            assert "layout_name" in sig.parameters

    def test_dispatcher_forwards_layout_name(self, mock_odoo_util, mock_log_util):
        from utility.util_autocad_dispatcher import UtilAutoCADDispatcher
        d = UtilAutoCADDispatcher(mock_odoo_util, mock_log_util, mode="com")
        d._active = MagicMock()
        d.draw_circle([0, 0], 20, "0", "配置2")
        d._active.draw_circle.assert_called_once_with([0, 0], 20, "0", "配置2")
