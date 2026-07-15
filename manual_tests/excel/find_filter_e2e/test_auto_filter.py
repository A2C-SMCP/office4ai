"""
Excel Find&Filter E2E — set:autoFilter + clear:autoFilter（自动筛选）

覆盖 ``excel:set:autoFilter``（按列条件应用）与 ``excel:clear:autoFilter``（清除）。

wire 形态（AddIn autoFilter.ts 确认）:
- set:autoFilter（address, criteria:[{columnIndex, filterOn, values?}], worksheetName?）→ ``{address}``
- clear:autoFilter（worksheetName?）→ **void**（无 data）

双重验证：openpyxl ``auto_filter_ref``——set 后 ref 落到筛选范围（如 'A1:C5'），clear 后 ref 清空。
含错误码两路径：address 空串（Zod min(1)）→ 4000；非法 address → 3009 RANGE_INVALID
（oasp#17 定案；Add-In 接线前（office-editor4ai#80）以 XFAIL 运行）。

运行方式:
    uv run python manual_tests/excel/find_filter_e2e/test_auto_filter.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.find_filter_e2e._fixtures import ensure_fixtures

FILT = "find_filter_e2e/filt.xlsx"

# 单列筛选：Region(col0) = East
_CRIT_ONE = [{"columnIndex": 0, "filterOn": "Values", "values": ["East"]}]
# 多列筛选：Region(col0)=East 且 Product(col1)=Apple
_CRIT_TWO = [
    {"columnIndex": 0, "filterOn": "Values", "values": ["East"]},
    {"columnIndex": 1, "filterOn": "Values", "values": ["Apple", "Banana"]},
]


def _addr_ok(data: dict[str, Any]) -> bool:
    addr = data.get("address")
    if not addr or "A1:C5" not in str(addr):
        print(f"   ❌ 响应 address 应含 'A1:C5'，实得 {addr!r}")
        return False
    return True


def validate_set_one(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    ref = reader.auto_filter_ref("Data")
    print(f"   set 后 auto_filter_ref={ref!r}")
    if ref != "A1:C5":
        print("   ❌ openpyxl auto_filter_ref 应为 'A1:C5'")
        return False
    print("   ✅ set:autoFilter 单列筛选已落地（ref=A1:C5）")
    return True


def validate_set_two(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    ref = reader.auto_filter_ref("Data")
    print(f"   多列 set 后 auto_filter_ref={ref!r}")
    if ref != "A1:C5":
        print("   ❌ openpyxl auto_filter_ref 应为 'A1:C5'")
        return False
    print("   ✅ set:autoFilter 多列筛选已落地（ref=A1:C5）")
    return True


def validate_clear(data: dict[str, Any], reader: WorkbookReader) -> bool:
    # clear 返回 void → 只读盘核对 ref 已清空。
    reader.reload()
    ref = reader.auto_filter_ref("Data")
    print(f"   clear 后 auto_filter_ref={ref!r}")
    if ref is not None:
        print("   ❌ clear 后 auto_filter_ref 应为 None")
        return False
    print("   ✅ clear:autoFilter 已清除筛选（ref=None）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="set 单列筛选",
        fixture_name=FILT,
        description="set:autoFilter A1:C5 Region=East → {address} + auto_filter_ref='A1:C5'",
        action="set:autoFilter",
        params={"address": "A1:C5", "criteria": _CRIT_ONE, "worksheet_name": "Data"},
        validator=validate_set_one,
        tags=["filter", "set"],
    ),
    ExcelCase(
        name="set 多列筛选",
        fixture_name=FILT,
        description="set:autoFilter A1:C5 Region=East & Product∈{Apple,Banana} → ref 落地",
        action="set:autoFilter",
        params={"address": "A1:C5", "criteria": _CRIT_TWO, "worksheet_name": "Data"},
        validator=validate_set_two,
        tags=["filter", "set", "multi"],
    ),
    ExcelCase(
        name="clear 清除筛选",
        fixture_name=FILT,
        description="先 set:autoFilter → clear:autoFilter → auto_filter_ref=None",
        action="clear:autoFilter",
        params={"worksheet_name": "Data"},
        pre_ops=[("set:autoFilter", {"address": "A1:C5", "criteria": _CRIT_ONE, "worksheet_name": "Data"})],
        validator=validate_clear,
        tags=["filter", "clear"],
    ),
    ExcelCase(
        name="错误码 4000 — address 空串（VALIDATION_ERROR）",
        fixture_name=FILT,
        description="set:autoFilter address='' → 4000（Zod min(1) 失败）",
        action="set:autoFilter",
        params={"address": "", "criteria": _CRIT_ONE, "worksheet_name": "Data"},
        expect_error_code="4000",
        tags=["error", "zod"],
    ),
    ExcelCase(
        name="错误码 3009 — 非法 address（RANGE_INVALID）",
        fixture_name=FILT,
        description="set:autoFilter address='ZZZZ99999999' → 3009（oasp#17）",
        action="set:autoFilter",
        params={"address": "ZZZZ99999999", "criteria": _CRIT_ONE, "worksheet_name": "Data"},
        expect_error_code="3009",
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Find&Filter E2E — set/clear:autoFilter", TEST_CASES)
