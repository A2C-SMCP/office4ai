"""
Excel PivotTable E2E — get:pivotTables（列举透视表）

覆盖 ``excel:get:pivotTables``。用 pre_op 先 insert 透视表，再 get 协议层回读核对。

wire 形态（AddIn pivotTable.ts 确认）:
- get:pivotTables（worksheetName?）→ ``{pivotTables:[{name, id}]}``（每条仅 2 字段）

含错误码：不存在的 worksheet → 3010 ELEMENT_NOT_FOUND(kind:worksheet)（oasp#17 定案；
Add-In 接线前 XFAIL）。

运行方式:
    uv run python manual_tests/excel/pivot_table_e2e/test_get_pivot_tables.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.pivot_table_e2e._fixtures import ensure_fixtures

PIVOT = "pivot_table_e2e/pivot.xlsx"


def validate_one(data: dict[str, Any]) -> bool:
    pts = data.get("pivotTables")
    if not isinstance(pts, list) or not pts:
        print(f"   ❌ pivotTables 非非空列表: {pts!r}")
        return False
    names = [p.get("name") for p in pts]
    print(f"   pivotTables={names}")
    if "P1" not in names:
        print("   ❌ 缺透视表 'P1'")
        return False
    for p in pts:
        missing = {"name", "id"} - set(p.keys())
        if missing:
            print(f"   ❌ 条目缺字段 {missing}: {p}")
            return False
    print("   ✅ get:pivotTables 回读单透视表（name/id）含 P1")
    return True


def validate_two(data: dict[str, Any]) -> bool:
    pts = data.get("pivotTables") or []
    names = sorted(str(p.get("name")) for p in pts)
    print(f"   pivotTables 数={len(pts)} names={names}")
    if not {"P1", "P2"}.issubset(set(names)):
        print("   ❌ 应同时含 P1 和 P2")
        return False
    print("   ✅ get:pivotTables 回读多透视表（P1+P2）")
    return True


def validate_empty(data: dict[str, Any]) -> bool:
    pts = data.get("pivotTables")
    print(f"   Blank pivotTables={pts!r}")
    if pts != []:
        print("   ❌ 空表应返回空 pivotTables 列表")
        return False
    print("   ✅ 无透视表 → 空列表")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="get 单透视表回读",
        fixture_name=PIVOT,
        description="先 insert 'P1' → get:pivotTables 回读 {name,id} 含 P1",
        action="get:pivotTables",
        params={"worksheet_name": "Data"},
        pre_ops=[
            (
                "insert:pivotTable",
                {"source_address": "A1:C5", "target_address": "E1", "name": "P1", "worksheet_name": "Data"},
            )
        ],
        validator=validate_one,
        tags=["get"],
    ),
    ExcelCase(
        name="get 多透视表",
        fixture_name=PIVOT,
        description="先 insert P1(E1)+P2(E20) → get:pivotTables 含 P1+P2",
        action="get:pivotTables",
        params={"worksheet_name": "Data"},
        pre_ops=[
            (
                "insert:pivotTable",
                {"source_address": "A1:C5", "target_address": "E1", "name": "P1", "worksheet_name": "Data"},
            ),
            (
                "insert:pivotTable",
                {"source_address": "A1:C5", "target_address": "E20", "name": "P2", "worksheet_name": "Data"},
            ),
        ],
        validator=validate_two,
        tags=["get", "multi"],
    ),
    ExcelCase(
        name="get 空表无透视表",
        fixture_name=PIVOT,
        description="get:pivotTables Blank（无透视表）→ 空列表",
        action="get:pivotTables",
        params={"worksheet_name": "Blank"},
        validator=validate_empty,
        tags=["get", "empty"],
    ),
    ExcelCase(
        name="错误码 3010 — 不存在的 worksheet（ELEMENT_NOT_FOUND kind:worksheet）",
        fixture_name=PIVOT,
        description="get:pivotTables worksheet_name='NoSuch' → 3010 + details.kind=worksheet（oasp#17）",
        action="get:pivotTables",
        params={"worksheet_name": "NoSuchSheet"},
        expect_error_code="3010",
        expect_error_details={"kind": "worksheet"},
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel PivotTable E2E — get:pivotTables", TEST_CASES)
