"""
Excel Format E2E — merge:cells + unmerge:cells（合并/取消合并单元格）

覆盖 ``excel:merge:cells``（``range.merge()`` 合为单格，保留左上值、其余清空）与
``excel:unmerge:cells``（``range.unmerge()`` 还原）。openpyxl ``reader.merged_ranges(sheet)``
读盘核对合并区域；``cell_value`` 核对左上保留 / 其余清空。

含错误码用例：非法 address → 3000 DOCUMENT_ERROR。

运行方式:
    uv run python manual_tests/excel/format_e2e/test_merge_cells.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.format_e2e._fixtures import ensure_fixtures

FMT = "format_e2e/fmt.xlsx"


def _addr_ok(data: dict[str, Any]) -> bool:
    if not data.get("address"):
        print(f"   ❌ 响应缺少 address: {data}")
        return False
    return True


def validate_merge(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    ranges = reader.merged_ranges("Merge")
    a1 = reader.cell_value("Merge", "A1")
    print(f"   merged_ranges={ranges}; A1={a1!r}")
    if "A1:C1" not in ranges:
        print("   ❌ merged_ranges 不含 A1:C1")
        return False
    if a1 != "M1":
        print("   ❌ 合并后左上值未保留为 'M1'")
        return False
    print("   ✅ A1:C1 已合并且左上值 'M1' 保留")
    return True


def validate_unmerge(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    ranges = reader.merged_ranges("Merge")
    print(f"   unmerge 后 merged_ranges={ranges}")
    if "A1:C1" in ranges:
        print("   ❌ 取消合并后 A1:C1 仍在 merged_ranges")
        return False
    print("   ✅ A1:C1 已取消合并（merged_ranges 不含该区域）")
    return True


def validate_merge_clears_others(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    b1 = reader.cell_value("Merge", "B1")
    c1 = reader.cell_value("Merge", "C1")
    print(f"   合并后 B1={b1!r} C1={c1!r}（预期 None）")
    if b1 is not None or c1 is not None:
        print("   ❌ 合并后非左上单元格应被清空")
        return False
    print("   ✅ 合并后 B1/C1 已清空（仅左上保留值）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="merge A1:C1",
        fixture_name=FMT,
        description="merge:cells Merge!A1:C1 → merged_ranges 含 A1:C1，左上 'M1' 保留",
        action="merge:cells",
        params={"address": "A1:C1", "worksheet_name": "Merge"},
        validator=validate_merge,
        tags=["merge"],
    ),
    ExcelCase(
        name="unmerge 还原",
        fixture_name=FMT,
        description="先 merge A1:C1 → unmerge:cells → merged_ranges 不含该区域",
        action="unmerge:cells",
        params={"address": "A1:C1", "worksheet_name": "Merge"},
        pre_ops=[("merge:cells", {"address": "A1:C1", "worksheet_name": "Merge"})],
        validator=validate_unmerge,
        tags=["unmerge"],
    ),
    ExcelCase(
        name="合并清空非左上单元格",
        fixture_name=FMT,
        description="merge A1:C1 后 B1/C1 应被清空（仅左上保留）",
        action="merge:cells",
        params={"address": "A1:C1", "worksheet_name": "Merge"},
        validator=validate_merge_clears_others,
        tags=["merge", "clear"],
    ),
    ExcelCase(
        name="错误码 3000 — 非法 address（DOCUMENT_ERROR）",
        fixture_name=FMT,
        description="merge:cells 传非法地址 → 3000（3009 为 dead code）",
        action="merge:cells",
        params={"address": "ZZZZ99999999", "worksheet_name": "Merge"},
        expect_error_code="3000",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Format E2E — merge/unmerge:cells", TEST_CASES)
