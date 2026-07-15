"""
Excel Read State E2E — get:worksheetInfo（工作表详细信息）

覆盖 ``excel:get:worksheetInfo``（可选 ``worksheetName``，省略=活动表）。验证 usedRange
行列数 / tableCount / chartCount，并用 openpyxl ``max_row``/``max_column`` 双重核对。
含错误码用例：ghost 表名 → ``3010 ELEMENT_NOT_FOUND``（details.kind=worksheet，oasp#17
定案；Add-In 接线前（office-editor4ai#80）以 XFAIL 运行）。

运行方式:
    uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test all
    uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test 6
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.read_state_e2e._fixtures import ensure_fixtures

MULTI = "read_state_e2e/multi_sheet.xlsx"


def _used(data: dict[str, Any]) -> dict[str, Any]:
    return data.get("usedRange") or {}


def validate_default_active(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name = data.get("name")
    print(f"   name={name!r}, usedRange={_used(data)}")
    if name != "Sheet1":
        print(f"   ⚠️  默认活动表为 {name!r}（夹具 activeSheet=Sheet1，若手动切过表会不同）")
    if "usedRange" not in data:
        print("   ❌ 缺少 usedRange")
        return False
    print(f"   ✅ 默认返回活动表 {name!r} 的 worksheetInfo")
    return True


def validate_named_sheet(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if data.get("name") != "Data":
        print(f"   ❌ name={data.get('name')!r}（预期 'Data'）")
        return False
    print(f"   ✅ worksheetName=Data 命中：name={data.get('name')!r}")
    return True


def validate_used_range_dims(data: dict[str, Any], reader: WorkbookReader) -> bool:
    used = _used(data)
    rc, cc = used.get("rowCount"), used.get("columnCount")
    reader.reload()
    ws = reader.wb["Data"]
    disk_rc, disk_cc = ws.max_row, ws.max_column
    print(f"   返回 usedRange={used}; openpyxl Data dims={disk_rc}x{disk_cc}")
    if rc != disk_rc or cc != disk_cc:
        print(f"   ❌ usedRange 行列 {rc}x{cc} 与磁盘 {disk_rc}x{disk_cc} 不符")
        return False
    print(f"   ✅ usedRange {rc} 行 × {cc} 列 与 openpyxl 一致")
    return True


def validate_counts(data: dict[str, Any], reader: WorkbookReader) -> bool:
    tc, cc = data.get("tableCount"), data.get("chartCount")
    print(f"   tableCount={tc}, chartCount={cc}")
    if not isinstance(tc, int) or not isinstance(cc, int):
        print("   ❌ tableCount/chartCount 非整数")
        return False
    if tc != 0 or cc != 0:
        print(f"   ⚠️  预填夹具应为 0 表 0 图（实际 {tc}/{cc}）")
    print(f"   ✅ tableCount={tc}, chartCount={cc} 字段就绪")
    return True


def validate_empty_used_range(data: dict[str, Any], reader: WorkbookReader) -> bool:
    used = _used(data)
    rc, cc = used.get("rowCount", 0), used.get("columnCount", 0)
    print(f"   空表 Report usedRange={used}")
    if (rc or 0) > 1 or (cc or 0) > 1:
        print(f"   ❌ 空表 usedRange 过大: {rc}x{cc}")
        return False
    print(f"   ✅ 空表 usedRange 极小（{rc}x{cc}）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="默认 active 表",
        fixture_name=MULTI,
        description="省略 worksheetName → 返回活动表 Sheet1 信息",
        action="get:worksheetInfo",
        validator=validate_default_active,
        tags=["basic"],
    ),
    ExcelCase(
        name="指定 worksheetName",
        fixture_name=MULTI,
        description="worksheetName=Data → 命中指定表",
        action="get:worksheetInfo",
        params={"worksheet_name": "Data"},
        validator=validate_named_sheet,
        tags=["named"],
    ),
    ExcelCase(
        name="usedRange 行列数",
        fixture_name=MULTI,
        description="Data 表 usedRange 4x3，与 openpyxl 核对",
        action="get:worksheetInfo",
        params={"worksheet_name": "Data"},
        validator=validate_used_range_dims,
        tags=["usedrange"],
    ),
    ExcelCase(
        name="tableCount/chartCount",
        fixture_name=MULTI,
        description="Data 表 tableCount=0 / chartCount=0 字段就绪",
        action="get:worksheetInfo",
        params={"worksheet_name": "Data"},
        validator=validate_counts,
        tags=["counts"],
    ),
    ExcelCase(
        name="空表 usedRange",
        fixture_name=MULTI,
        description="空表 Report 的 usedRange 极小",
        action="get:worksheetInfo",
        params={"worksheet_name": "Report"},
        validator=validate_empty_used_range,
        tags=["empty"],
    ),
    ExcelCase(
        # oasp#17 定案：工作表 not-found → 3010 ELEMENT_NOT_FOUND + details.kind=worksheet
        # （5xxx 块退役、不得降级 3000；旧 5001 与「真机 3000」口径均已过时，
        # 详见 docs/manual_tests/excel_e2e_dev_plan.md「错误码」）。
        name="错误码 3010 — ghost 表名（ELEMENT_NOT_FOUND kind:worksheet）",
        fixture_name=MULTI,
        description="worksheetName 传不存在的表名 → 3010 + details.kind=worksheet（oasp#17）",
        action="get:worksheetInfo",
        params={"worksheet_name": "GhostSheet_xyz"},
        expect_error_code="3010",
        expect_error_details={"kind": "worksheet"},
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Read State E2E — get:worksheetInfo", TEST_CASES)
