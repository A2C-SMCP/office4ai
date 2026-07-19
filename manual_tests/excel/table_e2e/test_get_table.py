"""
Excel Table E2E — get:table + get:tables（查表）

覆盖 ``excel:get:table``（富信息）与 ``excel:get:tables``（精简列表）。

wire 形态（AddIn table.ts 确认）:
- get:table（tableId, worksheetName?）→ ``TableInfo``：
  ``{name, id, address, rowCount, columnCount, columns:[{name,index}], styleName, showHeaders}``
- get:tables（worksheetName?）→ ``{tables:[{name, id, address}]}``（每条仅 3 字段）

读类事件，仅协议断言（夹具预置表已知形态）。含错误码：表不存在 → 3010
ELEMENT_NOT_FOUND(kind:table)（oasp#17 定案；Add-In 接线前 XFAIL）。

运行方式:
    uv run python manual_tests/excel/table_e2e/test_get_table.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.table_e2e._fixtures import ensure_fixtures

TBL = "table_e2e/tbl.xlsx"


def validate_get_sales(data: dict[str, Any]) -> bool:
    print(f"   keys={sorted(data.keys())}")
    required = ["name", "id", "address", "rowCount", "columnCount", "columns", "showHeaders"]
    missing = [k for k in required if k not in data]
    if missing:
        print(f"   ❌ TableInfo 缺字段: {missing}")
        return False
    if data["name"] != "SalesTable":
        print(f"   ❌ name 应为 'SalesTable'，实得 {data['name']!r}")
        return False
    cols = data.get("columns") or []
    col_names = [c.get("name") for c in cols]
    print(
        f"   rowCount={data['rowCount']} columnCount={data['columnCount']} columns={col_names} showHeaders={data['showHeaders']}"
    )
    if data["rowCount"] != 5 or data["columnCount"] != 3:
        print("   ❌ A1:C5 应为 rowCount=5 columnCount=3")
        return False
    if col_names != ["Name", "Score", "Age"]:
        print(f"   ❌ 列名应为 [Name,Score,Age]，实得 {col_names}")
        return False
    if not all(isinstance(c.get("index"), int) for c in cols):
        print("   ❌ 列缺少整数 index")
        return False
    print("   ✅ SalesTable 富信息完整（5×3，列名/index/showHeaders 正确）")
    return True


def validate_get_tables(data: dict[str, Any]) -> bool:
    tables = data.get("tables")
    if not isinstance(tables, list) or not tables:
        print(f"   ❌ tables 非非空列表: {tables!r}")
        return False
    names = [t.get("name") for t in tables]
    print(f"   Sales 表列表={names}")
    if "SalesTable" not in names:
        print("   ❌ 缺 SalesTable")
        return False
    for t in tables:
        missing = {"name", "id", "address"} - set(t.keys())
        if missing:
            print(f"   ❌ 条目缺字段 {missing}: {t}")
            return False
        if len(t.keys()) != 3:
            print(f"   ⚠️  条目字段非精简 3 项: {sorted(t.keys())}（get:tables 约定仅 name/id/address）")
    print("   ✅ get:tables 精简列表（name/id/address）含 SalesTable")
    return True


def validate_get_roster(data: dict[str, Any]) -> bool:
    if data.get("name") != "RosterTable":
        print(f"   ❌ name 应为 'RosterTable'，实得 {data.get('name')!r}")
        return False
    cols = [c.get("name") for c in (data.get("columns") or [])]
    print(f"   RosterTable rowCount={data.get('rowCount')} columns={cols}")
    if cols != ["Item", "Qty"]:
        print(f"   ❌ 列名应为 [Item,Qty]，实得 {cols}")
        return False
    print("   ✅ 跨表按 worksheet_name 取 RosterTable 成功")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="get:table 富信息",
        fixture_name=TBL,
        description="get:table SalesTable → TableInfo 完整（5×3，列名/index/showHeaders）",
        action="get:table",
        params={"table_id": "SalesTable", "worksheet_name": "Sales"},
        validator=validate_get_sales,
        tags=["get"],
    ),
    ExcelCase(
        name="get:tables 精简列表",
        fixture_name=TBL,
        description="get:tables Sales → [{name,id,address}] 含 SalesTable",
        action="get:tables",
        params={"worksheet_name": "Sales"},
        validator=validate_get_tables,
        tags=["get", "list"],
    ),
    ExcelCase(
        name="get:table 跨表按名取",
        fixture_name=TBL,
        description="get:table RosterTable（worksheet_name='Roster'）→ 列名 [Item,Qty]",
        action="get:table",
        params={"table_id": "RosterTable", "worksheet_name": "Roster"},
        validator=validate_get_roster,
        tags=["get"],
    ),
    ExcelCase(
        name="错误码 3010 — 表不存在（ELEMENT_NOT_FOUND kind:table）",
        fixture_name=TBL,
        description="get:table tableId='NoSuchTable' → 3010 + details.kind=table（oasp#17）",
        action="get:table",
        params={"table_id": "NoSuchTable", "worksheet_name": "Sales"},
        expect_error_code="3010",
        expect_error_details={"kind": "table"},
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Table E2E — get:table + get:tables", TEST_CASES)
