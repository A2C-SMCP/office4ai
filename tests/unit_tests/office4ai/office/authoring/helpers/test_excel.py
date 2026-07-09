"""fill_cells_lxml / locate_by_named_range 单测 | Unit tests for the Excel primitives."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from office4ai.office.authoring.helpers import (
    AnchorNotFoundError,
    fill_cells_lxml,
    locate_by_named_range,
)
from office4ai.office.authoring.helpers._ooxml import parse_xml, qn, read_package
from office4ai.office.authoring.helpers.excel import _parse_refers_to


def _sheet1_root(xlsx: Path):
    _, items = read_package(xlsx)
    return parse_xml(items["xl/worksheets/sheet1.xml"])


def _cell(root, ref: str):
    return root.find(f".//{qn('s:c')}[@r='{ref}']")


class TestFillCellsLxml:
    def test_preserves_chart_part(self, chart_xlsx: Path, tmp_path: Path) -> None:
        assert any("chart" in n for n in zipfile.ZipFile(chart_xlsx).namelist())
        out = tmp_path / "out.xlsx"
        fill_cells_lxml(chart_xlsx, out, {"数据": {"B2": 200}})
        assert any("chart" in n for n in zipfile.ZipFile(out).namelist()), "图表部件不应丢失"

    def test_writes_string_as_inline_str(self, chart_xlsx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.xlsx"
        n = fill_cells_lxml(chart_xlsx, out, {"数据": {"A2": "新收入"}})
        assert n == 1
        cell = _cell(_sheet1_root(out), "A2")
        assert cell is not None and cell.get("t") == "inlineStr"
        assert "新收入" in "".join(t.text or "" for t in cell.iter(qn("s:t")))

    def test_writes_number_as_value(self, chart_xlsx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.xlsx"
        fill_cells_lxml(chart_xlsx, out, {"数据": {"B2": 200}})
        cell = _cell(_sheet1_root(out), "B2")
        assert cell is not None and cell.get("t") is None  # 数字无 t 属性
        assert cell.find(qn("s:v")).text == "200"

    def test_none_clears_cell(self, chart_xlsx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.xlsx"
        fill_cells_lxml(chart_xlsx, out, {"数据": {"B2": None}})
        cell = _cell(_sheet1_root(out), "B2")
        assert cell is not None and len(list(cell)) == 0  # 无子元素 = 空

    def test_strips_formula_cache(self, chart_xlsx: Path, tmp_path: Path) -> None:
        # 原始模板里 B5 = "=B2-B3"（openpyxl 不写缓存值，这里主要验证 <f> 保留、无 <v>）
        out = tmp_path / "out.xlsx"
        fill_cells_lxml(chart_xlsx, out, {"数据": {"B2": 200}})
        b5 = _cell(_sheet1_root(out), "B5")
        assert b5 is not None
        assert b5.find(qn("s:f")) is not None, "公式应保留"
        assert b5.find(qn("s:v")) is None, "公式缓存值应被删除以逼重算"

    def test_in_place(self, chart_xlsx: Path) -> None:
        n = fill_cells_lxml(chart_xlsx, chart_xlsx, {"数据": {"A2": "原地"}})
        assert n == 1
        cell = _cell(_sheet1_root(chart_xlsx), "A2")
        assert "原地" in "".join(t.text or "" for t in cell.iter(qn("s:t")))

    def test_unknown_sheet_raises(self, chart_xlsx: Path, tmp_path: Path) -> None:
        with pytest.raises(AnchorNotFoundError):
            fill_cells_lxml(chart_xlsx, tmp_path / "o.xlsx", {"不存在的表": {"A1": 1}})

    def test_missing_ref_strict_raises(self, chart_xlsx: Path, tmp_path: Path) -> None:
        with pytest.raises(AnchorNotFoundError):
            fill_cells_lxml(chart_xlsx, tmp_path / "o.xlsx", {"数据": {"Z99": 1}}, strict=True)

    def test_missing_ref_non_strict_skips(self, chart_xlsx: Path, tmp_path: Path) -> None:
        out = tmp_path / "o.xlsx"
        n = fill_cells_lxml(chart_xlsx, out, {"数据": {"A2": "写入", "Z99": "跳过"}})
        assert n == 1  # 只有 A2 写入，Z99 不存在被跳过

    def test_writes_bool_as_boolean(self, chart_xlsx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.xlsx"
        fill_cells_lxml(chart_xlsx, out, {"数据": {"B2": True, "B3": False}})
        root = _sheet1_root(out)
        true_cell, false_cell = _cell(root, "B2"), _cell(root, "B3")
        assert true_cell.get("t") == "b" and true_cell.find(qn("s:v")).text == "1"
        assert false_cell.get("t") == "b" and false_cell.find(qn("s:v")).text == "0"

    def test_unsupported_value_type_raises(self, chart_xlsx: Path, tmp_path: Path) -> None:
        import datetime

        with pytest.raises(TypeError):
            fill_cells_lxml(chart_xlsx, tmp_path / "o.xlsx", {"数据": {"B2": datetime.date(2026, 1, 1)}})

    def test_sets_full_calc_on_load(self, chart_xlsx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.xlsx"
        fill_cells_lxml(chart_xlsx, out, {"数据": {"B2": 200}})
        _, items = read_package(out)
        workbook = parse_xml(items["xl/workbook.xml"])
        calc_pr = workbook.find(qn("s:calcPr"))
        assert calc_pr is not None and calc_pr.get("fullCalcOnLoad") == "1"


class TestLocateByNamedRange:
    def test_resolves_sheet_and_ref(self, chart_xlsx: Path) -> None:
        ranges = locate_by_named_range(chart_xlsx, "金额区")
        assert len(ranges) == 1
        nr = ranges[0]
        assert nr.sheet == "数据"
        assert nr.ref == "$B$2:$B$3"
        assert nr.local_sheet_id is None

    def test_unknown_name_returns_empty(self, chart_xlsx: Path) -> None:
        assert locate_by_named_range(chart_xlsx, "不存在") == []

    def test_sheet_scoped_defined_name(self, tmp_path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # 表级作用域定义名（localSheetId=0）
        ws.defined_names["表级名"] = DefinedName("表级名", attr_text="Sheet1!$A$1", localSheetId=0)
        path = tmp_path / "scoped.xlsx"
        wb.save(path)
        ranges = locate_by_named_range(path, "表级名")
        assert len(ranges) == 1
        assert ranges[0].local_sheet_id == 0
        assert ranges[0].sheet == "Sheet1"


class TestParseRefersTo:
    def test_plain_sheet_and_ref(self) -> None:
        assert _parse_refers_to("数据!$B$2:$C$3") == ("数据", "$B$2:$C$3")

    def test_quoted_sheet_name_with_escaped_quote(self) -> None:
        # OOXML 用 '' 转义单引号：'It''s'!$A$1 -> sheet = "It's"
        assert _parse_refers_to("'It''s'!$A$1") == ("It's", "$A$1")

    def test_no_bang_returns_ref_only(self) -> None:
        assert _parse_refers_to("$A$1:$B$2") == (None, "$A$1:$B$2")
