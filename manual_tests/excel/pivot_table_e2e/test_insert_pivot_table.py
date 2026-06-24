"""
Excel PivotTable E2E — insert:pivotTable（创建透视表）

覆盖 ``excel:insert:pivotTable``：指定 name / 默认 name。仅用 sourceAddress + targetAddress
创建空透视表骨架（不构造 rows/columns/values/filters）。

wire 形态（AddIn pivotTable.ts 确认）:
- insert:pivotTable（sourceAddress, targetAddress, name?, worksheetName?）→ ``{name}``
  （name 省略时默认 "PivotTable"）。source 与 target 须同一 worksheet。

主验证：协议返回非空 name（透视表落地由 ``test_get_pivot_tables.py`` 协议回读核对）。
含错误码两路径：非法 sourceAddress → 3000；sourceAddress 空串（Zod min(1)）→ 4000。

运行方式:
    uv run python manual_tests/excel/pivot_table_e2e/test_insert_pivot_table.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.pivot_table_e2e._fixtures import ensure_fixtures

PIVOT = "pivot_table_e2e/pivot.xlsx"


def validate_named(data: dict[str, Any]) -> bool:
    name = data.get("name")
    print(f"   insert 返回: name={name!r}")
    if name != "SalesPivot":
        print(f"   ❌ name 应为 'SalesPivot'，实得 {name!r}")
        return False
    print("   ✅ 指定 name 创建透视表 'SalesPivot'")
    return True


def validate_default(data: dict[str, Any]) -> bool:
    name = data.get("name")
    print(f"   insert（无 name）返回: name={name!r}")
    if not name:
        print("   ❌ 默认命名应返回非空 name")
        return False
    print(f"   ✅ 默认命名创建透视表（name={name!r}）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="insert 指定 name",
        fixture_name=PIVOT,
        description="insert:pivotTable source=A1:C5 target=E1 name='SalesPivot' → {name:'SalesPivot'}",
        action="insert:pivotTable",
        params={"source_address": "A1:C5", "target_address": "E1", "name": "SalesPivot", "worksheet_name": "Data"},
        validator=validate_named,
        tags=["insert"],
    ),
    ExcelCase(
        name="insert 默认 name",
        fixture_name=PIVOT,
        description="insert:pivotTable source=A1:C5 target=E1（省略 name）→ 非空 name",
        action="insert:pivotTable",
        params={"source_address": "A1:C5", "target_address": "E1", "worksheet_name": "Data"},
        validator=validate_default,
        tags=["insert", "default"],
    ),
    ExcelCase(
        name="错误码 3000 — 非法 sourceAddress（DOCUMENT_ERROR）",
        fixture_name=PIVOT,
        description="insert:pivotTable sourceAddress='ZZZZ99999999' → 3000",
        action="insert:pivotTable",
        params={"source_address": "ZZZZ99999999", "target_address": "E1", "worksheet_name": "Data"},
        expect_error_code="3000",
        tags=["error"],
    ),
    ExcelCase(
        name="错误码 4000 — sourceAddress 空串（VALIDATION_ERROR）",
        fixture_name=PIVOT,
        description="insert:pivotTable sourceAddress='' → 4000（Zod min(1) 失败）",
        action="insert:pivotTable",
        params={"source_address": "", "target_address": "E1", "worksheet_name": "Data"},
        expect_error_code="4000",
        tags=["error", "zod"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel PivotTable E2E — insert:pivotTable", TEST_CASES)
