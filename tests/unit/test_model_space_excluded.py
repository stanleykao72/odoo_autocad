# -*- coding: utf-8 -*-
"""
Model 空間排除測試。

Odoo 流程的標題欄與 TABLE 都放在配置的圖紙空間，Model 不在處理範圍。
另有實務理由：某些圖面的 Model 內含 proxy／未載入物件，透過 COM 列舉會使
AutoCAD 2014 直接崩潰（實測 E169A-203.dwg 兩次），更不應去掃它。
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
    return util


def _logged(backend):
    return " ".join(c.args[0] for c in backend.log.safe_log_insert.call_args_list)


class TestResolveWorkingLayout:
    def test_uses_active_layout_when_not_model(self, backend):
        backend.get_active_layout = MagicMock(return_value="E169A-203")
        backend.get_doc_layouts = MagicMock()

        assert backend._resolve_working_layout() == "E169A-203"
        backend.get_doc_layouts.assert_not_called()

    def test_falls_back_to_first_layout_when_active_is_model(self, backend):
        backend.get_active_layout = MagicMock(return_value="Model")
        backend.get_doc_layouts = MagicMock(return_value=["配置1", "配置2"])

        assert backend._resolve_working_layout() == "配置1"
        assert "Model" in _logged(backend), "改用其他配置時必須留下紀錄"

    def test_returns_none_when_only_model_exists(self, backend):
        backend.get_active_layout = MagicMock(return_value="Model")
        backend.get_doc_layouts = MagicMock(return_value=[])

        assert backend._resolve_working_layout() is None
        assert "沒有任何配置" in _logged(backend)


class TestProcessPrNoSkipsModel:
    def test_model_is_skipped_without_touching_layout(self, backend):
        backend.get_layout_from_name = MagicMock()

        backend.process_pr_no("Model")

        backend.get_layout_from_name.assert_not_called(), "不可取用 Model 的實體"
        assert "不在處理範圍" in _logged(backend)


class TestGetSingleLayoutValuesSkipsModel:
    def test_model_returns_empty_without_enumerating(self, backend):
        backend.get_layout_from_name = MagicMock()

        assert backend.get_single_layout_values("Model") == {}
        backend.get_layout_from_name.assert_not_called()


class TestClearTableIdSkipsModel:
    def test_refuses_when_active_layout_is_model(self, backend):
        backend.get_active_layout = MagicMock(return_value="Model")
        backend.get_layout_from_name = MagicMock()
        backend.get_layout_table_block = MagicMock()

        backend.clear_table_id()          # 未指定配置 -> 回退到作用中配置

        backend.get_layout_table_block.assert_not_called(), \
            "破壞性操作不可對 Model 執行"
        assert "請先切換到要清除的配置" in _logged(backend)

    def test_explicit_layout_still_works(self, backend):
        backend.get_layout_from_name = MagicMock(return_value=MagicMock())
        backend.get_layout_table_block = MagicMock(return_value=[])

        backend.clear_table_id("配置1")

        backend.get_layout_table_block.assert_called_once()


class TestAlreadyExcludedPaths:
    """既有已排除 Model 的路徑，確保不退化"""

    def test_get_doc_layouts_excludes_model(self, backend):
        model = MagicMock(); model.Name = "Model"
        lay = MagicMock(); lay.Name = "配置1"
        backend.doc.Layouts = [model, lay]

        assert backend.get_doc_layouts() == ["配置1"]

    def test_get_block_attributes_returns_empty_for_model(self, backend):
        layout = MagicMock(); layout.Name = "Model"
        backend.acad.ActiveDocument.ActiveLayout = layout

        assert backend.get_block_attributes() == {}


class TestScanElementsNotSilentlyEmpty:
    """COM 端未實作 scan_elements —— 不可回傳空結果假裝成功

    舊版回傳 {"elements": [], "summary": {"total_count": 0, ...}}，呼叫端
    無法分辨「圖面真的沒有元素」與「這個功能根本沒做」。
    """

    def test_raises_not_implemented(self, backend):
        with pytest.raises(NotImplementedError):
            backend.scan_elements()

    def test_message_points_to_ipc_mode(self, backend):
        with pytest.raises(NotImplementedError, match="IPC"):
            backend.scan_elements()

    def test_still_satisfies_abc_contract(self, backend):
        """必須仍是 UtilAutoCAD 的方法，否則 ABC 會讓類別無法實例化"""
        assert callable(backend.scan_elements)


class TestFakeStubsRemoved:
    """已刪除的假資料 stub 不可復活 —— 真實對應為 get_product / import2boq"""

    @pytest.mark.parametrize("name", ["search_products", "push_boq_data"])
    def test_fake_odoo_stub_is_gone(self, name):
        from utility.util_odoo import UtilOdoo
        assert not hasattr(UtilOdoo, name), (
            f"{name} 是回傳假資料的死碼，真實對應為 "
            "get_product() / import2boq()")

    def test_real_counterparts_exist(self):
        from utility.util_odoo import UtilOdoo
        for name in ("get_product", "get_setup", "get_color",
                     "import2boq", "boq2pr"):
            assert callable(getattr(UtilOdoo, name))

    def test_scan_entities_is_gone(self, backend):
        assert not hasattr(backend, "scan_entities")
