"""
Excel PivotTable E2E — delete:pivotTable（删除透视表）

覆盖 ``excel:delete:pivotTable``。用 ``ExcelCase.flow`` 多步流：先 insert（可指定 name），
再 delete，最后 get:pivotTables 协议层回读核对消失。

wire 形态（AddIn pivotTable.ts 确认）:
- delete:pivotTable（pivotTableName, worksheetName?）→ **void**（无 data）

含错误码：delete 不存在的 pivotTableName → 3000 DOCUMENT_ERROR。

运行方式:
    uv run python manual_tests/excel/pivot_table_e2e/test_delete_pivot_table.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.pivot_table_e2e._fixtures import ensure_fixtures
from manual_tests.excel.test_helpers import excel_op
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

PIVOT = "pivot_table_e2e/pivot.xlsx"


async def _names(workspace: OfficeWorkspace, uri: str) -> list[Any] | None:
    ok, data, err = await excel_op(workspace, uri, "get:pivotTables", worksheet_name="Data")
    if not ok:
        print(f"   ❌ get:pivotTables 回读失败: {err}")
        return None
    return [p.get("name") for p in (data or {}).get("pivotTables", [])]


async def _insert(workspace: OfficeWorkspace, uri: str, name: str, target: str) -> bool:
    ok, _, err = await excel_op(
        workspace,
        uri,
        "insert:pivotTable",
        source_address="A1:C5",
        target_address=target,
        name=name,
        worksheet_name="Data",
    )
    if not ok:
        print(f"   ❌ 预置 insert:pivotTable {name!r} 失败: {err}")
    return ok


async def _flow_delete_one(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    # 插两个，删其一，验证「删的没了、留的还在」。
    if not (await _insert(workspace, uri, "P1", "E1") and await _insert(workspace, uri, "P2", "E20")):
        return False
    before = await _names(workspace, uri)
    print(f"   插入后 pivotTables={before}（删除目标=P1）")
    if before is None or "P1" not in before or "P2" not in before:
        print("   ❌ 预置透视表未就位")
        return False
    ok, data, err = await excel_op(workspace, uri, "delete:pivotTable", pivot_table_name="P1", worksheet_name="Data")
    if not ok:
        print(f"   ❌ delete:pivotTable 失败: {err}")
        return False
    print(f"   delete 返回 data={data}（应为 void）")
    after = await _names(workspace, uri)
    print(f"   删除后 pivotTables={after}")
    if after is None or "P1" in after:
        print("   ❌ 目标 P1 仍在")
        return False
    if "P2" not in after:
        print("   ❌ 非目标 P2 被误删")
        return False
    print("   ✅ delete:pivotTable 删除目标、保留其它")
    return True


async def _flow_delete_last(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    # 删唯一透视表，验证 get 回到空列表。
    if not await _insert(workspace, uri, "Solo", "E1"):
        return False
    ok, _, err = await excel_op(workspace, uri, "delete:pivotTable", pivot_table_name="Solo", worksheet_name="Data")
    if not ok:
        print(f"   ❌ delete:pivotTable 失败: {err}")
        return False
    after = await _names(workspace, uri)
    print(f"   删除唯一透视表后 pivotTables={after}")
    if after != []:
        print("   ❌ 删除唯一透视表后应为空列表")
        return False
    print("   ✅ delete:pivotTable 删除唯一透视表 → 空列表")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="delete 其一保留其它",
        fixture_name=PIVOT,
        description="insert P1+P2→delete P1→get:pivotTables 验证 P1 没了、P2 还在",
        flow=_flow_delete_one,
        tags=["delete"],
    ),
    ExcelCase(
        name="delete 唯一透视表→空",
        fixture_name=PIVOT,
        description="insert Solo→delete Solo→get:pivotTables 回到空列表",
        flow=_flow_delete_last,
        tags=["delete", "empty"],
    ),
    ExcelCase(
        name="错误码 3000 — delete 不存在的透视表（DOCUMENT_ERROR）",
        fixture_name=PIVOT,
        description="delete:pivotTable pivotTableName='NoSuch' → 3000",
        action="delete:pivotTable",
        params={"pivot_table_name": "NoSuchPivot", "worksheet_name": "Data"},
        expect_error_code="3000",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel PivotTable E2E — delete:pivotTable", TEST_CASES)
