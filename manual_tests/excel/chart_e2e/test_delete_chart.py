"""
Excel Chart E2E — delete:chart（删除图表）

覆盖 ``excel:delete:chart``。因图表名由 Excel 自动生成，用 ``ExcelCase.flow`` 多步流：
先 insert 拿 name，再 delete，最后 get:charts 协议层回读核对图表已消失。

wire 形态（AddIn chart.ts 确认）:
- delete:chart（chartName, worksheetName?）→ **void**（无 data）

含错误码：delete 不存在的 chartName → 3000 DOCUMENT_ERROR。

运行方式:
    uv run python manual_tests/excel/chart_e2e/test_delete_chart.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.chart_e2e._fixtures import ensure_fixtures
from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.test_helpers import excel_op
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

CHART = "chart_e2e/chart.xlsx"


async def _names(workspace: OfficeWorkspace, uri: str) -> list[Any] | None:
    ok, data, err = await excel_op(workspace, uri, "get:charts", worksheet_name="Data")
    if not ok:
        print(f"   ❌ get:charts 回读失败: {err}")
        return None
    return [c.get("name") for c in (data or {}).get("charts", [])]


async def _flow_delete_one(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    # 插两个图表，删其一，验证「删的没了、留的还在」。
    ok, d1, _ = await excel_op(
        workspace, uri, "insert:chart", source_address="A1:C4", chart_type="ColumnClustered", worksheet_name="Data"
    )
    ok2, d2, _ = await excel_op(
        workspace, uri, "insert:chart", source_address="A1:C4", chart_type="Line", worksheet_name="Data"
    )
    if not (ok and ok2 and (d1 or {}).get("name") and (d2 or {}).get("name")):
        print("   ❌ 预置 insert:chart 失败")
        return False
    target, keep = d1["name"], d2["name"]
    before = await _names(workspace, uri)
    print(f"   插入后 charts={before}（删除目标={target!r}）")
    if before is None or target not in before or keep not in before:
        print("   ❌ 预置图表未就位")
        return False
    ok, data, err = await excel_op(workspace, uri, "delete:chart", chart_name=target, worksheet_name="Data")
    if not ok:
        print(f"   ❌ delete:chart 失败: {err}")
        return False
    print(f"   delete 返回 data={data}（应为 void）")
    after = await _names(workspace, uri)
    print(f"   删除后 charts={after}")
    if after is None or target in after:
        print(f"   ❌ 目标图表 {target!r} 仍在")
        return False
    if keep not in after:
        print(f"   ❌ 非目标图表 {keep!r} 被误删")
        return False
    print("   ✅ delete:chart 删除目标、保留其它")
    return True


async def _flow_delete_last(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    # 删唯一图表，验证 get:charts 回到空列表。
    ok, d, err = await excel_op(
        workspace, uri, "insert:chart", source_address="A1:C4", chart_type="Pie", worksheet_name="Data"
    )
    if not ok or not (d or {}).get("name"):
        print(f"   ❌ 预置 insert:chart 失败: {err}")
        return False
    name = d["name"]
    ok, _, err = await excel_op(workspace, uri, "delete:chart", chart_name=name, worksheet_name="Data")
    if not ok:
        print(f"   ❌ delete:chart 失败: {err}")
        return False
    after = await _names(workspace, uri)
    print(f"   删除唯一图表后 charts={after}")
    if after != []:
        print("   ❌ 删除唯一图表后应为空列表")
        return False
    print("   ✅ delete:chart 删除唯一图表 → 空列表")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="delete 其一保留其它",
        fixture_name=CHART,
        description="insert×2→delete 其一→get:charts 验证删的没了、留的还在",
        flow=_flow_delete_one,
        tags=["delete"],
    ),
    ExcelCase(
        name="delete 唯一图表→空",
        fixture_name=CHART,
        description="insert→delete→get:charts 回到空列表",
        flow=_flow_delete_last,
        tags=["delete", "empty"],
    ),
    ExcelCase(
        name="错误码 3000 — delete 不存在的图表（DOCUMENT_ERROR）",
        fixture_name=CHART,
        description="delete:chart chartName='NoSuch' → 3000",
        action="delete:chart",
        params={"chart_name": "NoSuchChart", "worksheet_name": "Data"},
        expect_error_code="3000",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Chart E2E — delete:chart", TEST_CASES)
