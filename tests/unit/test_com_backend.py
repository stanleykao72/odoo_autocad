# -*- coding: utf-8 -*-
"""
Tests for COM backend (UtilAutoCAD).

Verifies:
- get_active_layout() returns str, not COM object
- get_doc_layouts() returns List[str], not COM objects
- get_layouts_values() returns {"all": [...]}
- get_layouts_header_id_to_pr() returns {"all": [...]}
- get_block_attributes() returns dict
- get_single_layout_values() returns dict with correct keys
- clear_table_id() accepts string layout name
"""
import pytest
import sys
from unittest.mock import MagicMock, Mock, patch, PropertyMock


# Patch the COM handles that util_autocad already bound at import time.
#
# 早期版本只在 sys.modules 尚無 'pythoncom' 時才注入 MagicMock，因此單獨執行
# 本檔會通過、整套執行時（其他測試已先 import 真正的 pythoncom）卻失敗。
# 這裡改為直接 patch util_autocad 模組內的名稱，與 import 順序無關。
@pytest.fixture(autouse=True)
def mock_com_handles():
    import utility.util_autocad as ua
    with patch.object(ua, 'pythoncom', MagicMock()) as mock_pythoncom, \
            patch.object(ua, 'client', MagicMock()) as mock_client:
        # com_error 必須是可被 except 捕捉的真實例外類別
        mock_pythoncom.com_error = type('com_error', (Exception,), {})
        yield mock_pythoncom, mock_client


def _make_mock_layout(name, blocks=None, has_block=True):
    """Create a mock COM Layout object"""
    layout = MagicMock()
    layout.Name = name
    if has_block and blocks is not None:
        layout.Block = blocks
    elif not has_block:
        layout.Block = None
    return layout


def _make_mock_attribute(tag, value):
    att = MagicMock()
    att.TagString = tag
    att.TextString = value
    return att


def _make_mock_block_ref(name, attributes=None, is_attr_block=False):
    """Create a mock AcDbBlockReference"""
    block = MagicMock()
    block.ObjectName = "AcDbBlockReference"
    block.Name = name
    block.EffectiveName = name
    block.HasAttributes = bool(attributes)
    if attributes:
        block.GetAttributes.return_value = attributes
    return block


def _make_mock_table(rows_data, header_id="H001"):
    """Create a mock AcDbTable with row data.
    rows_data: list of lists, each inner list is [pos, product_no, w, h, len, thick, qty, desc, detail_id]
    """
    table = MagicMock()
    table.ObjectName = "AcDbTable"
    # Row 0 = header, Row 1 = column names, Row 2+ = data
    total_rows = 2 + len(rows_data)
    table.Rows = total_rows
    table.Columns = 9

    def get_cell_value(row, col):
        if row == 0 and col == 8:
            return header_id
        if row >= 2:
            data_idx = row - 2
            if data_idx < len(rows_data):
                return rows_data[data_idx][col] if col < len(rows_data[data_idx]) else ""
        return ""

    table.GetCellValue = get_cell_value
    return table


@pytest.fixture
def com_backend(mock_odoo_util, mock_log_util):
    from utility.util_autocad import UtilAutoCAD
    util = UtilAutoCAD(mock_odoo_util, mock_log_util)
    return util


class TestCOMGetActiveLayout:
    def test_returns_string(self, com_backend):
        """get_active_layout() must return a string, not a COM object"""
        mock_doc = MagicMock()
        mock_layout = MagicMock()
        mock_layout.Name = "Layout1"
        mock_doc.ActiveLayout = mock_layout
        com_backend.doc = mock_doc

        result = com_backend.get_active_layout()
        assert isinstance(result, str)
        assert result == "Layout1"

    def test_returns_none_when_no_doc(self, com_backend):
        com_backend.doc = None
        result = com_backend.get_active_layout()
        assert result is None


class TestCOMGetDocLayouts:
    def test_returns_list_of_strings(self, com_backend):
        """get_doc_layouts() must return List[str], not COM objects"""
        layout1 = MagicMock()
        layout1.Name = "Layout1"
        layout2 = MagicMock()
        layout2.Name = "Layout2"
        model = MagicMock()
        model.Name = "Model"

        mock_doc = MagicMock()
        mock_doc.Layouts = [model, layout1, layout2]
        com_backend.doc = mock_doc

        result = com_backend.get_doc_layouts()
        assert isinstance(result, list)
        assert all(isinstance(name, str) for name in result)
        assert result == ["Layout1", "Layout2"]

    def test_excludes_model(self, com_backend):
        model = MagicMock()
        model.Name = "Model"
        mock_doc = MagicMock()
        mock_doc.Layouts = [model]
        com_backend.doc = mock_doc

        result = com_backend.get_doc_layouts()
        assert result == []

    def test_returns_empty_list_on_error(self, com_backend):
        """Must return [] not None on error"""
        com_backend.doc = MagicMock()
        com_backend.doc.Layouts = None

        result = com_backend.get_doc_layouts()
        assert result == []


class TestCOMGetLayoutsValues:
    def test_returns_dict_with_all_key(self, com_backend):
        """get_layouts_values() must return {"all": [...]}"""
        # Setup: one layout with attribute block + table
        attr_block = _make_mock_block_ref("title_block", [
            _make_mock_attribute("project_name", "TestProject"),
            _make_mock_attribute("pr_no", "PR001"),
        ], is_attr_block=True)
        pr_block = _make_mock_block_ref("pr_no")
        table = _make_mock_table([
            ["1", "PROD-A", "100", "200", "300", "10", "5", "desc", ""]
        ])
        table.ObjectName = "AcDbTable"

        layout_obj = _make_mock_layout("Layout1", blocks=[attr_block, pr_block, table])

        mock_doc = MagicMock()
        layout1_mock = MagicMock()
        layout1_mock.Name = "Layout1"
        model_mock = MagicMock()
        model_mock.Name = "Model"
        mock_doc.Layouts = [model_mock, layout1_mock]

        com_backend.doc = mock_doc

        # Mock get_layout_from_name to return our prepared layout
        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)
        # Mock internal methods
        com_backend.get_layout_attribute_blocks_value = MagicMock(return_value={
            "project_name": "TestProject", "pr_no": "PR001"
        })
        com_backend.get_layout_table_block = MagicMock(return_value=[table])
        com_backend.get_table_data = MagicMock(return_value=("H001", [{"product_no": "PROD-A"}]))

        result = com_backend.get_layouts_values()
        assert isinstance(result, dict)
        assert "all" in result
        assert isinstance(result["all"], list)
        assert len(result["all"]) == 1
        assert result["all"][0]["layout_name"] == "Layout1"

    def test_returns_empty_dict_when_no_layouts(self, com_backend):
        mock_doc = MagicMock()
        mock_doc.Layouts = []
        com_backend.doc = mock_doc

        result = com_backend.get_layouts_values()
        assert isinstance(result, dict)


class TestCOMGetSingleLayoutValues:
    def test_returns_dict_with_expected_keys(self, com_backend):
        """get_single_layout_values() returns dict with layout_name, header_id, detail"""
        layout_obj = _make_mock_layout("TestLayout")

        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)
        com_backend.get_layout_attribute_blocks_value = MagicMock(return_value={
            "project_name": "P1", "pr_no": "PR001"
        })
        table = _make_mock_table([["1", "PROD-A", "", "", "", "", "5", "", ""]])
        com_backend.get_layout_table_block = MagicMock(return_value=[table])
        com_backend.get_table_data = MagicMock(return_value=("H001", [{"product_no": "PROD-A", "qty": "5"}]))

        result = com_backend.get_single_layout_values("TestLayout")
        assert isinstance(result, dict)
        assert result["layout_name"] == "TestLayout"
        assert result["header_id"] == "H001"
        assert isinstance(result["detail"], list)

    def test_returns_empty_dict_for_nonexistent_layout(self, com_backend):
        com_backend.get_layout_from_name = MagicMock(return_value=None)

        result = com_backend.get_single_layout_values("NoSuchLayout")
        assert result == {}


class TestCOMGetBlockAttributes:
    def test_returns_dict(self, com_backend):
        """get_block_attributes() returns dict of tag:value"""
        mock_layout = MagicMock()
        mock_layout.Name = "Layout1"

        mock_doc = MagicMock()
        mock_doc.ActiveLayout = mock_layout
        com_backend.acad = MagicMock()
        com_backend.acad.ActiveDocument = mock_doc

        attr_block = _make_mock_block_ref("title", [
            _make_mock_attribute("project_name", "MyProject"),
            _make_mock_attribute("job_working_plan_name", "Plan1"),
        ])
        com_backend.get_attribute_block = MagicMock(return_value=attr_block)

        pr_block = _make_mock_block_ref("pr_no")
        com_backend.get_attribute_block_with_name = MagicMock(return_value=pr_block)
        com_backend.get_block_text = MagicMock(return_value="PR001")

        result = com_backend.get_block_attributes()
        assert isinstance(result, dict)
        assert result["project_name"] == "MyProject"
        assert result["pr_no"] == "PR001"
        assert result["layout_name"] == "Layout1"

    def test_returns_empty_dict_when_not_connected(self, com_backend):
        com_backend.acad = None
        result = com_backend.get_block_attributes()
        assert result == {}


class TestCOMGetLayoutsHeaderIdToPr:
    def test_returns_dict_with_all_key(self, com_backend):
        """get_layouts_header_id_to_pr() returns {"all": [header_id, ...]}"""
        layout_obj = _make_mock_layout("Layout1")

        mock_doc = MagicMock()
        l1 = MagicMock(); l1.Name = "Layout1"
        model = MagicMock(); model.Name = "Model"
        mock_doc.Layouts = [model, l1]
        com_backend.doc = mock_doc

        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)
        table = _make_mock_table([["1", "PROD-A", "", "", "", "", "5", "", ""]])
        com_backend.get_layout_table_block = MagicMock(return_value=[table])
        com_backend.get_table_data = MagicMock(return_value=("H001", [{"product_no": "PROD-A"}]))

        result = com_backend.get_layouts_header_id_to_pr()
        assert isinstance(result, dict)
        assert "all" in result
        assert isinstance(result["all"], list)
        assert "H001" in result["all"]

    def test_layout_without_block_does_not_reuse_previous_header_id(self, com_backend):
        """迴歸測試：沒有 Block 的 layout 不可沿用上一個 layout 的 header_id。

        舊版 header_id / detail_list 在迴圈外殘留，導致同一個 header_id
        被重複收集（或第一個 layout 就沒 Block 時 UnboundLocalError）。
        """
        with_block = _make_mock_layout("Layout1")
        without_block = _make_mock_layout("Layout2", has_block=True)
        without_block.Block = None  # 沒有可處理的 Block

        mock_doc = MagicMock()
        l1 = MagicMock(); l1.Name = "Layout1"
        l2 = MagicMock(); l2.Name = "Layout2"
        mock_doc.Layouts = [l1, l2]
        com_backend.doc = mock_doc

        com_backend.get_layout_from_name = MagicMock(
            side_effect=lambda n: with_block if n == "Layout1" else without_block)
        com_backend.get_layout_table_block = MagicMock(return_value=[MagicMock()])
        com_backend.get_table_data = MagicMock(return_value=("H001", [{"product_no": "A"}]))

        result = com_backend.get_layouts_header_id_to_pr()
        # 只有 Layout1 有資料 — H001 不可出現兩次
        assert result["all"] == ["H001"]

    def test_first_layout_without_block_does_not_raise(self, com_backend):
        """第一個 layout 就沒有 Block 時不可 UnboundLocalError"""
        no_block = _make_mock_layout("Layout1")
        no_block.Block = None

        mock_doc = MagicMock()
        l1 = MagicMock(); l1.Name = "Layout1"
        mock_doc.Layouts = [l1]
        com_backend.doc = mock_doc
        com_backend.get_layout_from_name = MagicMock(return_value=no_block)

        result = com_backend.get_layouts_header_id_to_pr()
        assert result == {"all": []}


class TestCOMGetTableData:
    def test_empty_table_list_returns_none_header(self, com_backend):
        """迴歸測試：table_list 為空時 header_id 未定義會 UnboundLocalError"""
        header_id, detail_list = com_backend.get_table_data([])
        assert header_id is None
        assert detail_list == []


class TestCOMGetAttributeBlock:
    def test_returns_none_when_no_matching_block(self, com_backend):
        """迴歸測試：找不到屬性塊時 return_block 未初始化會 UnboundLocalError，
        導致真正的原因被外層 except 吞掉並記錄成誤導的錯誤訊息。"""
        other = _make_mock_block_ref("other", [_make_mock_attribute("foo", "bar")])
        layout = _make_mock_layout("Layout1", blocks=[other])

        result = com_backend.get_attribute_block(layout)
        assert result is None
        # 必須是「未找到屬性塊」而不是例外訊息
        logged = " ".join(str(c) for c in com_backend.log.safe_log_insert.call_args_list)
        assert "未找到屬性塊" in logged
        assert "UnboundLocalError" not in logged

    def test_returns_block_when_matching_tag_present(self, com_backend):
        target = _make_mock_block_ref("title", [
            _make_mock_attribute("project_name", "P1")])
        layout = _make_mock_layout("Layout1", blocks=[target])

        result = com_backend.get_attribute_block(layout)
        assert result is target


class TestCOMClearTableId:
    def test_accepts_string_layout_name(self, com_backend):
        """clear_table_id() should accept a string layout name"""
        layout_obj = _make_mock_layout("Layout1")
        table = _make_mock_table([["1", "P1", "", "", "", "", "1", "", "D1"]])

        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)
        com_backend.get_layout_table_block = MagicMock(return_value=[table])
        com_backend.set_table_value = MagicMock()

        # Should not raise
        com_backend.clear_table_id("Layout1")
        com_backend.get_layout_from_name.assert_called_with("Layout1")

    def test_uses_active_layout_when_none(self, com_backend):
        """clear_table_id(None) uses get_active_layout()"""
        mock_doc = MagicMock()
        mock_layout = MagicMock()
        mock_layout.Name = "ActiveLayout"
        mock_doc.ActiveLayout = mock_layout
        com_backend.doc = mock_doc

        layout_obj = _make_mock_layout("ActiveLayout")
        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)
        com_backend.get_layout_table_block = MagicMock(return_value=[])

        com_backend.clear_table_id()
        com_backend.get_layout_from_name.assert_called_with("ActiveLayout")

    def test_clear_all_tables_id(self, com_backend):
        """clear_all_tables_id() iterates all layout names"""
        mock_doc = MagicMock()
        l1 = MagicMock(); l1.Name = "L1"
        l2 = MagicMock(); l2.Name = "L2"
        model = MagicMock(); model.Name = "Model"
        mock_doc.Layouts = [model, l1, l2]
        com_backend.doc = mock_doc

        layout_obj = _make_mock_layout("L1")
        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)
        com_backend.get_layout_table_block = MagicMock(return_value=[])

        com_backend.clear_all_tables_id()
        # Should have been called with string names
        calls = com_backend.get_layout_from_name.call_args_list
        called_names = [c[0][0] for c in calls]
        assert "L1" in called_names
        assert "L2" in called_names


class TestCOMSetBlockAttributes:
    def test_returns_true_on_success(self, com_backend):
        """set_block_attributes() is implemented and returns True"""
        mock_doc = MagicMock()
        mock_layout = MagicMock()
        mock_layout.Name = "Layout1"

        attrs = [
            _make_mock_attribute("PROJECT_NAME", "Old"),
            _make_mock_attribute("PR_NO", "Old"),
        ]
        block = MagicMock()
        block.ObjectName = "AcDbBlockReference"
        block.HasAttributes = True
        block.GetAttributes.return_value = attrs

        mock_layout.Block = [block]
        mock_doc.Layouts.Item.return_value = mock_layout
        com_backend.acad = MagicMock()
        com_backend.acad.ActiveDocument = mock_doc

        result = com_backend.set_block_attributes({"project_name": "New"}, "Layout1")
        assert result is True


class TestCOMProcessPrNo:
    """Test process_pr_no receives string layout name, resolves to COM object internally.

    This is the exact bug scenario: connect_autocad() calls get_active_layout()
    which returns a string, then passes it to process_pr_no(). If process_pr_no
    doesn't resolve the string to a COM layout object, it crashes with
    "'str' object has no attribute 'Block'".
    """

    def test_process_pr_no_accepts_string(self, com_backend):
        """process_pr_no(layout_name) must accept a string and resolve it"""
        layout_obj = _make_mock_layout("E172A-201")
        com_backend.get_layout_from_name = MagicMock(return_value=layout_obj)

        pr_block = _make_mock_block_ref("pr_no")
        com_backend.get_attribute_block_with_name = MagicMock(return_value=pr_block)
        com_backend.get_block_text = MagicMock(return_value="PR001")
        com_backend.odoo_util = MagicMock()
        com_backend.odoo_util.get_project.return_value = {
            'id': 42, 'name': 'Test Project',
            'job_working_plan_id': 1, 'job_working_plan_name': 'Plan A'
        }
        attr_block = _make_mock_block_ref("title_block")
        com_backend.get_attribute_block = MagicMock(return_value=attr_block)
        com_backend.set_attribute_value = MagicMock()

        # Must NOT raise "'str' object has no attribute 'Block'"
        com_backend.process_pr_no("E172A-201")

        # Verify it resolved the string to COM layout
        com_backend.get_layout_from_name.assert_called_with("E172A-201")
        # Verify it passed the COM layout object (not the string) to internal methods
        com_backend.get_attribute_block_with_name.assert_called_with(layout_obj, 'pr_no')
        com_backend.get_attribute_block.assert_called_with(layout_obj)
        assert com_backend.pr_no == "PR001"
        assert com_backend.project_id == 42

    def test_process_pr_no_handles_missing_layout(self, com_backend):
        """process_pr_no with nonexistent layout should not crash"""
        com_backend.get_layout_from_name = MagicMock(return_value=None)
        # Must NOT raise
        com_backend.process_pr_no("NonExistent")

    def test_connect_autocad_passes_string_to_process_pr_no(self, com_backend):
        """The full chain: connect_autocad → get_active_layout(str) → process_pr_no(str)

        This test catches the exact bug: if get_active_layout returns a string
        but process_pr_no expects a COM layout object, it would crash with
        "'str' object has no attribute 'Block'".
        """
        # Mock get_active_layout to return a string (which it does now)
        com_backend.get_active_layout = MagicMock(return_value="Layout1")
        com_backend.process_pr_no = MagicMock()
        com_backend.clear_main_body = MagicMock()

        # Mock the COM connection parts
        mock_doc = MagicMock()
        mock_doc.Name = "test.dwg"
        mock_doc.FullName = "C:\\test.dwg"
        com_backend.acad = MagicMock()
        com_backend.doc = mock_doc

        main_body = MagicMock()
        com_backend.connect_autocad(main_body)

        # process_pr_no must be called with a STRING, not a COM object
        com_backend.process_pr_no.assert_called_once()
        arg = com_backend.process_pr_no.call_args[0][0]
        assert isinstance(arg, str), f"Expected str, got {type(arg).__name__}: {arg}"
        assert arg == "Layout1"
