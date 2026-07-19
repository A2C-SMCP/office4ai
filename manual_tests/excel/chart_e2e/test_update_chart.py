"""
Excel Chart E2E — update:chart（更新图表属性）

覆盖 ``excel:update:chart`` 的 title / chartType / position 更新。

因图表名由 Excel 自动生成，用 ``ExcelCase.flow`` 多步流：先 insert 拿 name，再 update，
最后 get:charts 协议层回读核对更新已生效。

wire 形态（AddIn chart.ts 确认）:
- update:chart（chartName, properties{title?,chartType?,sourceAddress?,position?}, worksheetName?）
  → ``{name}``

含错误码：update 不存在的 chartName → 3010 ELEMENT_NOT_FOUND(kind:chart)（oasp#17 定案；
Add-In 接线前（office-editor4ai#80）以 XFAIL 运行）。

运行方式:
    uv run python manual_tests/excel/chart_e2e/test_update_chart.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.chart_e2e._fixtures import ensure_fixtures
from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.test_helpers import excel_op
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

CHART = "chart_e2e/chart.xlsx"


async def _insert(workspace: OfficeWorkspace, uri: str, **kw: Any) -> str | None:
    kw.setdefault("source_address", "A1:C4")
    kw.setdefault("worksheet_name", "Data")
    ok, data, err = await excel_op(workspace, uri, "insert:chart", **kw)
    if not ok or not (data or {}).get("name"):
        print(f"   ❌ 预置 insert:chart 失败: {err}")
        return None
    return data["name"]


async def _find_chart(workspace: OfficeWorkspace, uri: str, name: str) -> dict[str, Any] | None:
    ok, data, err = await excel_op(workspace, uri, "get:charts", worksheet_name="Data")
    if not ok:
        print(f"   ❌ get:charts 回读失败: {err}")
        return None
    for c in (data or {}).get("charts", []):
        if c.get("name") == name:
            return c
    print(f"   ❌ get:charts 未找到图表 {name!r}")
    return None


async def _flow_update_title(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    name = await _insert(workspace, uri, chart_type="ColumnClustered", title="Before")
    if not name:
        return False
    print(f"   插入图表 name={name!r} title='Before'")
    ok, _, err = await excel_op(
        workspace, uri, "update:chart", chart_name=name, properties={"title": "After"}, worksheet_name="Data"
    )
    if not ok:
        print(f"   ❌ update:chart 失败: {err}")
        return False
    c = await _find_chart(workspace, uri, name)
    if not c:
        return False
    print(f"   回读 title={c.get('title')!r}")
    if c.get("title") != "After":
        print("   ❌ title 未更新为 'After'")
        return False
    print("   ✅ update title：'Before' → 'After' 已生效")
    return True


async def _flow_update_type(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    name = await _insert(workspace, uri, chart_type="ColumnClustered", title="T")
    if not name:
        return False
    print(f"   插入图表 name={name!r} type=ColumnClustered")
    ok, _, err = await excel_op(
        workspace, uri, "update:chart", chart_name=name, properties={"chartType": "Line"}, worksheet_name="Data"
    )
    if not ok:
        print(f"   ❌ update:chart 失败: {err}")
        return False
    c = await _find_chart(workspace, uri, name)
    if not c:
        return False
    print(f"   回读 chartType={c.get('chartType')!r}")
    if "Line" not in str(c.get("chartType")):
        print("   ❌ chartType 未更新为 Line")
        return False
    print("   ✅ update chartType：ColumnClustered → Line 已生效")
    return True


async def _flow_update_position(workspace: OfficeWorkspace, uri: str, reader: WorkbookReader) -> bool:
    name = await _insert(workspace, uri, chart_type="ColumnClustered", title="P", position={"top": 10, "left": 10})
    if not name:
        return False
    ok, _, err = await excel_op(
        workspace,
        uri,
        "update:chart",
        chart_name=name,
        properties={"position": {"top": 250, "left": 320}},
        worksheet_name="Data",
    )
    if not ok:
        print(f"   ❌ update:chart 失败: {err}")
        return False
    c = await _find_chart(workspace, uri, name)
    if not c:
        return False
    print(f"   回读 top={c.get('top')} left={c.get('left')}")
    if abs(float(c.get("top", 0)) - 250) > 2 or abs(float(c.get("left", 0)) - 320) > 2:
        print("   ❌ position 未更新到 (top=250,left=320)")
        return False
    print("   ✅ update position：(10,10) → (250,320) 已生效")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="update title",
        fixture_name=CHART,
        description="insert→update title 'Before'→'After'→get:charts 回读核对",
        flow=_flow_update_title,
        tags=["update", "title"],
    ),
    ExcelCase(
        name="update chartType",
        fixture_name=CHART,
        description="insert ColumnClustered→update chartType=Line→回读核对",
        flow=_flow_update_type,
        tags=["update", "type"],
    ),
    ExcelCase(
        name="update position",
        fixture_name=CHART,
        description="insert→update position(top=250,left=320)→回读核对",
        flow=_flow_update_position,
        tags=["update", "position"],
    ),
    ExcelCase(
        name="错误码 3010 — update 不存在的图表（ELEMENT_NOT_FOUND kind:chart）",
        fixture_name=CHART,
        description="update:chart chartName='NoSuch' → 3010 + details.kind=chart（oasp#17）",
        action="update:chart",
        params={"chart_name": "NoSuchChart", "properties": {"title": "x"}, "worksheet_name": "Data"},
        expect_error_code="3010",
        expect_error_details={"kind": "chart"},
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Chart E2E — update:chart", TEST_CASES)
