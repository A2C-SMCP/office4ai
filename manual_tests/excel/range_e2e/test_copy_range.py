"""
Excel Range E2E — copy:range（复制范围到目标位置）

覆盖 ``excel:copy:range``（sourceAddress → targetAddress，targetAddress 为目标左上角）。
spec 未定义粘贴选项（值/格式始终整体复制）。在 Data 表内把已有数据复制到空白区域，
openpyxl 读盘核对目标区域内容与源一致、源区域不变。

运行方式:
    uv run python manual_tests/excel/range_e2e/test_copy_range.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.range_e2e._fixtures import ensure_fixtures

GRID = "range_e2e/grid.xlsx"


def _addr_ok(data: dict[str, Any]) -> bool:
    if not data.get("address"):
        print(f"   ❌ 响应缺少 address: {data}")
        return False
    return True


def validate_copy_row(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a6, b6, c6 = (reader.cell_value("Data", a) for a in ("A6", "B6", "C6"))
    a1 = reader.cell_value("Data", "A1")  # 源不变
    print(f"   目标 A6:C6 = {(a6, b6, c6)}；源 A1={a1!r}")
    if (a6, b6, c6) != ("Region", "Q1", "Q2"):
        print("   ❌ 目标行内容与源不一致")
        return False
    if a1 != "Region":
        print("   ❌ 源区域被破坏")
        return False
    print("   ✅ 表头行 A1:C1 → A6:C6 复制成功，源不变")
    return True


def validate_copy_block(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    e2 = reader.cell_value("Data", "E2")  # 源 A2='North'
    g4 = reader.cell_value("Data", "G4")  # 源 C4=350
    print(f"   目标 E2={e2!r} G4={g4!r}（源 A2='North' / C4=350）")
    if e2 != "North" or g4 != 350:
        print("   ❌ 块复制内容错位/缺失")
        return False
    print("   ✅ 数据块 A2:C4 → E2:G4 复制成功")
    return True


def validate_copy_cell(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    e1 = reader.cell_value("Data", "E1")
    print(f"   目标 E1={e1!r}（源 A1='Region'）")
    if e1 != "Region":
        print("   ❌ 单元格复制失败")
        return False
    print("   ✅ 单元格 A1 → E1 复制成功")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="复制单行到空白区",
        fixture_name=GRID,
        description="copy Data!A1:C1 → A6（空白行），目标内容与源一致、源不变",
        action="copy:range",
        params={"source_address": "A1:C1", "target_address": "A6", "worksheet_name": "Data"},
        validator=validate_copy_row,
        tags=["row"],
    ),
    ExcelCase(
        name="复制数据块到空白区",
        fixture_name=GRID,
        description="copy Data!A2:C4 → E2，目标 E2='North' / G4=350",
        action="copy:range",
        params={"source_address": "A2:C4", "target_address": "E2", "worksheet_name": "Data"},
        validator=validate_copy_block,
        tags=["block"],
    ),
    ExcelCase(
        name="复制单元格",
        fixture_name=GRID,
        description="copy Data!A1 → E1，目标 E1='Region'",
        action="copy:range",
        params={"source_address": "A1", "target_address": "E1", "worksheet_name": "Data"},
        validator=validate_copy_cell,
        tags=["cell"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Range E2E — copy:range", TEST_CASES)
