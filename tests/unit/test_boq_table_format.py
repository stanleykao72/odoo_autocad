# -*- coding: utf-8 -*-
"""
BOQ 表格讀取規則測試（新版表格格式）。

實測 2026.07.31-TEST.dwg 的表格：Rows=5, Columns=12

  r0: c0='加工細節' ... c7='HEADER_ID' c8=(header_id 值) c9~c11=''
  r1: c0='位置' c1='加工編號' c2='寬W' c3='高H' c4='長L' c5='厚度'
      c6='數量' c7='備註' c8='DETAIL ID'
  r2: 資料列
  r3: 空白列
  r4: c0='合計'                       <- 合計列

規則：
  1. 只讀 c0~c8；DETAIL ID 右側欄位數不固定，一律忽略
  2. 合計列不進 Odoo
  3. 表格判定需「欄數 >= 9」而非「剛好 9」——
     舊版寫死 cols == 9，這張 12 欄的表格會整個被跳過
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
    return UtilAutoCAD(mock_odoo_util, mock_log_util)


def _table(grid):
    """grid: list of rows, each a list of cell values"""
    t = MagicMock()
    t.ObjectName = "AcDbTable"
    t.Rows = len(grid)
    t.Columns = max(len(r) for r in grid)

    def get(r, c):
        row = grid[r]
        return row[c] if c < len(row) else ""
    t.GetCellValue.side_effect = get
    return t


def _real_grid():
    """與實機圖面相同的 12 欄表格"""
    return [
        ['加工細節', '', '', '', '', '', '', 'HEADER_ID', 'H999', '', '', ''],
        ['位置', '加工編號', '寬W', '高H', '長L', '厚度', '數量', '備註',
         'DETAIL ID', '', '', ''],
        ['R1F', 'E169-209', '88.4', '', '2918.5', '', '4', '蓋板', '',
         'X9', 'X10', 'X11'],
        ['', '', '', '', '', '', '', '', '', '', '', ''],
        ['合計', '', '', '', '', '', '99', '', '', '', '', ''],
    ]


class TestTableDetection:
    def test_accepts_more_than_nine_columns(self, backend):
        """迴歸測試：舊版寫死 cols == 9，12 欄的表格會整個被跳過"""
        assert backend.chk_legal_table(_table(_real_grid())) == "Y"

    def test_accepts_exactly_nine_columns(self, backend):
        grid = [r[:9] for r in _real_grid()]
        assert backend.chk_legal_table(_table(grid)) == "Y"

    def test_rejects_fewer_than_nine_columns(self, backend):
        grid = [r[:8] for r in _real_grid()]
        assert backend.chk_legal_table(_table(grid)) == "N"

    def test_rejects_when_header_id_label_missing(self, backend):
        grid = _real_grid()
        grid[0][7] = '其他標題'
        assert backend.chk_legal_table(_table(grid)) == "N"


class TestReadRange:
    def test_only_reads_columns_0_to_8(self, backend):
        _, detail = backend.get_table_data([_table(_real_grid())])
        assert len(detail) == 1
        assert set(detail[0].keys()) == set(backend.BOQ_COLUMNS)
        # DETAIL ID 右側的 X9/X10/X11 不可出現在任何值中
        assert not {'X9', 'X10', 'X11'} & set(map(str, detail[0].values()))

    def test_header_id_read_from_row0_col8(self, backend):
        header_id, _ = backend.get_table_data([_table(_real_grid())])
        assert header_id == 'H999'

    def test_column_mapping(self, backend):
        _, detail = backend.get_table_data([_table(_real_grid())])
        row = detail[0]
        assert row['position'] == 'R1F'
        assert row['product_no'] == 'E169-209'
        assert row['width'] == '88.4'
        assert row['len'] == '2918.5'
        assert row['qty'] == '4'
        assert row['desc'] == '蓋板'


class TestTotalRowExcluded:
    def test_total_row_not_in_detail(self, backend):
        _, detail = backend.get_table_data([_table(_real_grid())])
        assert all(r['position'] != '合計' for r in detail)

    def test_total_row_excluded_even_with_quantity(self, backend):
        """合計列的數量欄有值時（本例 99）也不可被當成資料列"""
        _, detail = backend.get_table_data([_table(_real_grid())])
        assert all(r['qty'] != '99' for r in detail)

    @pytest.mark.parametrize("label", ['合計', '總計', '小計', 'Total', ' 合計 '])
    def test_total_label_variants(self, backend, label):
        grid = _real_grid()
        grid[4][0] = label
        _, detail = backend.get_table_data([_table(grid)])
        assert len(detail) == 1, f"{label!r} 應被視為合計列"

    def test_normal_position_is_not_treated_as_total(self, backend):
        grid = _real_grid()
        grid[4][0] = 'R2F'
        grid[4][1] = 'E169-300'
        _, detail = backend.get_table_data([_table(grid)])
        assert len(detail) == 2, "一般位置名稱不可被誤判為合計列"


class TestWriteBackSkipsTotalRow:
    def test_detail_id_not_written_to_total_row(self, backend):
        table = _table(_real_grid())
        backend.set_table_value = MagicMock()
        layout = MagicMock()
        layout.Block = [table]
        backend.get_layout_from_name = MagicMock(return_value=layout)
        backend.get_layout_table_block = MagicMock(return_value=[table])

        backend.set_layouts_tables_id([{
            'layout_name': '01', 'header_id': 'H999',
            'detail': [{'product_no': 'E169-209', 'detail_id': 'D1'}],
        }])

        written_rows = [c.args[1] for c in backend.set_table_value.call_args_list]
        assert 4 not in written_rows, "合計列（r4）不可寫入 detail_id"
        assert 2 in written_rows, "資料列（r2）應寫入"


class TestHeaderIdColumnIsScanned:
    """HEADER_ID 位置改為掃描第 0 列，不再寫死欄索引。

    第 0 列的標題（如「加工細節」）是跨欄合併的，合併範圍一調整
    HEADER_ID 就會落到不同欄；寫死 (0,7) 會因版面微調而失效。
    """

    def _grid_with_header_at(self, col, total_cols=12):
        row0 = [''] * total_cols
        row0[0] = '加工細節'
        row0[col] = 'HEADER_ID'
        row0[col + 1] = 'H777'
        rest = _real_grid()[1:]
        return [row0] + [r + [''] * (total_cols - len(r)) for r in rest]

    @pytest.mark.parametrize("col", [5, 7, 9])
    def test_detects_header_id_at_any_column(self, backend, col):
        table = _table(self._grid_with_header_at(col))
        assert backend.chk_legal_table(table) == "Y"
        header_id, _ = backend.get_table_data([table])
        assert header_id == 'H777', f"HEADER_ID 在 c{col} 時應讀到右一格的值"

    def test_case_and_whitespace_tolerant(self, backend):
        grid = _real_grid()
        grid[0][7] = '  header_id  '
        table = _table(grid)
        assert backend.chk_legal_table(table) == "Y"
        header_id, _ = backend.get_table_data([table])
        assert header_id == 'H999'

    def test_rejected_when_label_is_last_column(self, backend):
        """標籤在最後一欄 —— 右邊沒有存值的格子"""
        grid = _real_grid()
        grid[0][7] = ''
        for r in grid:
            while len(r) > 12:
                r.pop()
        grid[0][11] = 'HEADER_ID'
        assert backend.chk_legal_table(_table(grid)) == "N"

    def test_write_back_uses_scanned_column(self, backend):
        table = _table(self._grid_with_header_at(9))
        backend.set_table_value = MagicMock()
        layout = MagicMock()
        layout.Block = [table]
        backend.get_layout_from_name = MagicMock(return_value=layout)
        backend.get_layout_table_block = MagicMock(return_value=[table])

        backend.set_layouts_tables_id([{
            'layout_name': '01', 'header_id': 'NEW',
            'detail': [{'product_no': 'E169-209', 'detail_id': 'D1'}],
        }])

        header_writes = [c.args for c in backend.set_table_value.call_args_list
                         if c.args[1] == 0]
        assert header_writes, "應寫入 header_id"
        assert header_writes[0][2] == 10, "HEADER_ID 在 c9 時，值應寫到 c10"
