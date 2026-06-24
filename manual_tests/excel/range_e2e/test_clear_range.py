"""
Excel Range E2E — clear:range（清除内容/格式/全部）

覆盖 ``excel:clear:range`` 的三种 ``clearType``（contents / formats / all）。先用
``set:range`` 铺数据（pre_ops），再清除（被测动作），最后 openpyxl 读盘核对：
- contents/all：单元格值被清空（None）；
- formats：值**保留**（只清格式不动内容）；
- 子区域清除：目标范围清空，相邻范围不受影响。

运行方式:
    uv run python manual_tests/excel/range_e2e/test_clear_range.py --test all
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


def validate_clear_contents(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a1 = reader.cell_value("Blank", "A1")
    b2 = reader.cell_value("Blank", "B2")
    print(f"   清除后 Blank A1={a1!r} B2={b2!r}")
    if a1 is not None or b2 is not None:
        print("   ❌ clearType=contents 后单元格仍有值")
        return False
    print("   ✅ contents 已清空（A1:B2 全空）")
    return True


def validate_clear_all(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a1 = reader.cell_value("Blank", "A1")
    print(f"   清除后 Blank A1={a1!r}")
    if a1 is not None:
        print("   ❌ clearType=all 后 A1 仍有值")
        return False
    print("   ✅ all 已清空")
    return True


def validate_clear_formats_keeps_value(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a1 = reader.cell_value("Blank", "A1")
    print(f"   清格式后 Blank A1={a1!r}（预期保留 'stay'）")
    if a1 != "stay":
        print("   ❌ clearType=formats 不应清除内容，但值丢失")
        return False
    print("   ✅ formats 只清格式，内容 'stay' 保留")
    return True


def validate_clear_subrange(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    b2 = reader.cell_value("Data", "B2")  # 被清
    a2 = reader.cell_value("Data", "A2")  # 相邻，应保留 'North'
    b3 = reader.cell_value("Data", "B3")  # 相邻行，应保留 200
    print(f"   Data B2={b2!r}（清）, A2={a2!r}, B3={b3!r}（相邻应保留）")
    if b2 is not None:
        print("   ❌ 目标 B2:C2 未清空")
        return False
    if a2 != "North" or b3 != 200:
        print("   ❌ 相邻单元格被误清")
        return False
    print("   ✅ 仅 B2:C2 清空，相邻 A2/B3 完好")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="clear contents",
        fixture_name=GRID,
        description="填 A1:B2 后 clearType=contents → 值清空",
        action="clear:range",
        params={"address": "A1:B2", "clear_type": "contents", "worksheet_name": "Blank"},
        pre_ops=[("set:range", {"address": "A1:B2", "values": "x", "worksheet_name": "Blank"})],
        validator=validate_clear_contents,
        tags=["contents"],
    ),
    ExcelCase(
        name="clear all",
        fixture_name=GRID,
        description="填 A1:B2 后 clearType=all → 值清空",
        action="clear:range",
        params={"address": "A1:B2", "clear_type": "all", "worksheet_name": "Blank"},
        pre_ops=[("set:range", {"address": "A1:B2", "values": "x", "worksheet_name": "Blank"})],
        validator=validate_clear_all,
        tags=["all"],
    ),
    ExcelCase(
        name="clear formats 保留内容",
        fixture_name=GRID,
        description="填 A1='stay' 后 clearType=formats → 内容保留（只清格式）",
        action="clear:range",
        params={"address": "A1", "clear_type": "formats", "worksheet_name": "Blank"},
        pre_ops=[("set:range", {"address": "A1", "values": "stay", "worksheet_name": "Blank"})],
        validator=validate_clear_formats_keeps_value,
        tags=["formats"],
    ),
    ExcelCase(
        name="清子区域不影响相邻",
        fixture_name=GRID,
        description="clear Data!B2:C2 contents → 仅该区域清空，A2/B3 不受影响",
        action="clear:range",
        params={"address": "B2:C2", "clear_type": "contents", "worksheet_name": "Data"},
        validator=validate_clear_subrange,
        tags=["subrange"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Range E2E — clear:range", TEST_CASES)
