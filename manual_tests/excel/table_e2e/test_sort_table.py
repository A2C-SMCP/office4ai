"""
Excel Table E2E — sort:table（表格排序）

覆盖 ``excel:sort:table``：单列升 / 单列降 / 多级排序（断 tie）。

wire 形态（AddIn table.ts 确认）:
- sort:table（tableId, sortFields:[{columnIndex, ascending?}], worksheetName?）→ **void**
  （ascending 省略 = true）。

双重验证：openpyxl 读 SalesTable 正文 Name 列（A2:A5）核对排序后行序。
SalesTable 正文：Bob(85,30) / Alice(85,25) / Dave(78,35) / Carol(88,28)。
列索引 0=Name 1=Score 2=Age；Age 唯一供单列排序，Score 含并列 85 供多级断 tie。

含错误码：表不存在 → 3010 ELEMENT_NOT_FOUND(kind:table)（oasp#17 定案；Add-In 接线前 XFAIL）。

运行方式:
    uv run python manual_tests/excel/table_e2e/test_sort_table.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.table_e2e._fixtures import ensure_fixtures

TBL = "table_e2e/tbl.xlsx"


def _names(reader: WorkbookReader) -> list[Any]:
    reader.reload()
    return [reader.cell_value("Sales", f"A{r}") for r in range(2, 6)]


def _check(reader: WorkbookReader, expected: list[str], label: str) -> bool:
    got = _names(reader)
    print(f"   排序后 Name 列={got}（期望 {expected}）")
    if got != expected:
        print(f"   ❌ {label} 行序不符")
        return False
    print(f"   ✅ {label} 行序正确")
    return True


def validate_sort_age_asc(data: dict[str, Any], reader: WorkbookReader) -> bool:
    # Age 升：25,28,30,35 → Alice,Carol,Bob,Dave
    return _check(reader, ["Alice", "Carol", "Bob", "Dave"], "Age 单列升序")


def validate_sort_age_desc(data: dict[str, Any], reader: WorkbookReader) -> bool:
    # Age 降：35,30,28,25 → Dave,Bob,Carol,Alice
    return _check(reader, ["Dave", "Bob", "Carol", "Alice"], "Age 单列降序")


def validate_sort_multi(data: dict[str, Any], reader: WorkbookReader) -> bool:
    # Score↑ 再 Age↑：Dave(78) → [85: Alice(25) 先于 Bob(30)] → Carol(88)
    return _check(reader, ["Dave", "Alice", "Bob", "Carol"], "Score↑→Age↑ 多级")


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="sort 单列升序（Age）",
        fixture_name=TBL,
        description="sort:table SalesTable by Age(col2) asc → Alice,Carol,Bob,Dave",
        action="sort:table",
        params={
            "table_id": "SalesTable",
            "sort_fields": [{"columnIndex": 2, "ascending": True}],
            "worksheet_name": "Sales",
        },
        validator=validate_sort_age_asc,
        tags=["sort", "asc"],
    ),
    ExcelCase(
        name="sort 单列降序（Age）",
        fixture_name=TBL,
        description="sort:table SalesTable by Age(col2) desc → Dave,Bob,Carol,Alice",
        action="sort:table",
        params={
            "table_id": "SalesTable",
            "sort_fields": [{"columnIndex": 2, "ascending": False}],
            "worksheet_name": "Sales",
        },
        validator=validate_sort_age_desc,
        tags=["sort", "desc"],
    ),
    ExcelCase(
        name="sort 多级（Score↑→Age↑ 断 tie）",
        fixture_name=TBL,
        description="sort:table by Score(col1)↑ then Age(col2)↑ → Dave,Alice,Bob,Carol",
        action="sort:table",
        params={
            "table_id": "SalesTable",
            "sort_fields": [{"columnIndex": 1, "ascending": True}, {"columnIndex": 2, "ascending": True}],
            "worksheet_name": "Sales",
        },
        validator=validate_sort_multi,
        tags=["sort", "multi"],
    ),
    ExcelCase(
        name="错误码 3010 — 表不存在（ELEMENT_NOT_FOUND kind:table）",
        fixture_name=TBL,
        description="sort:table tableId='NoSuchTable' → 3010 + details.kind=table（oasp#17）",
        action="sort:table",
        params={"table_id": "NoSuchTable", "sort_fields": [{"columnIndex": 0}], "worksheet_name": "Sales"},
        expect_error_code="3010",
        expect_error_details={"kind": "table"},
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Table E2E — sort:table", TEST_CASES)
