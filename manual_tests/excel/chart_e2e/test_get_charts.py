"""
Excel Chart E2E — get:charts（列举图表）

覆盖 ``excel:get:charts``。用 pre_op 先 insert 图表，再 get:charts 协议层回读核对
（图表名由 Excel 自动生成，故按 chartType/title 断言，不按 name）。

wire 形态（AddIn chart.ts 确认）:
- get:charts（worksheetName?）→ ``{charts:[{name, chartType, title, top, left, width, height}]}``

含错误码：不存在的 worksheet → 3010 ELEMENT_NOT_FOUND(kind:worksheet)（oasp#17 定案；
Add-In 接线前（office-editor4ai#80）以 XFAIL 运行）。

运行方式:
    uv run python manual_tests/excel/chart_e2e/test_get_charts.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.chart_e2e._fixtures import ensure_fixtures
from manual_tests.excel.e2e_case import ExcelCase, run_main

CHART = "chart_e2e/chart.xlsx"


def validate_one_chart(data: dict[str, Any]) -> bool:
    charts = data.get("charts")
    if not isinstance(charts, list) or not charts:
        print(f"   ❌ charts 非非空列表: {charts!r}")
        return False
    c = charts[0]
    print(f"   chart[0]: name={c.get('name')!r} chartType={c.get('chartType')!r} title={c.get('title')!r}")
    required = {"name", "chartType", "title", "top", "left", "width", "height"}
    missing = required - set(c.keys())
    if missing:
        print(f"   ❌ 条目缺字段: {missing}")
        return False
    if "Column" not in str(c.get("chartType")):
        print(f"   ❌ chartType 应含 'Column'，实得 {c.get('chartType')!r}")
        return False
    if c.get("title") != "Col Chart":
        print(f"   ❌ title 应为 'Col Chart'，实得 {c.get('title')!r}")
        return False
    if not all(isinstance(c.get(k), (int, float)) for k in ("top", "left", "width", "height")):
        print("   ❌ 位置字段应为数值")
        return False
    print("   ✅ get:charts 回读单图表完整（chartType/title/位置）")
    return True


def validate_two_charts(data: dict[str, Any]) -> bool:
    charts = data.get("charts") or []
    types = [c.get("chartType") for c in charts]
    print(f"   charts 数={len(charts)} types={types}")
    if len(charts) < 2:
        print("   ❌ 应至少 2 个图表")
        return False
    print("   ✅ get:charts 回读多图表（≥2）")
    return True


def validate_empty(data: dict[str, Any]) -> bool:
    charts = data.get("charts")
    print(f"   Blank charts={charts!r}")
    if charts != []:
        print("   ❌ 空表应返回空 charts 列表")
        return False
    print("   ✅ 无图表 → 空列表")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="get:charts 单图表回读",
        fixture_name=CHART,
        description="先 insert ColumnClustered 'Col Chart' → get:charts 回读 chartType/title/位置",
        action="get:charts",
        params={"worksheet_name": "Data"},
        pre_ops=[
            (
                "insert:chart",
                {
                    "source_address": "A1:C4",
                    "chart_type": "ColumnClustered",
                    "title": "Col Chart",
                    "worksheet_name": "Data",
                },
            )
        ],
        validator=validate_one_chart,
        tags=["get"],
    ),
    ExcelCase(
        name="get:charts 多图表",
        fixture_name=CHART,
        description="先 insert 2 个图表 → get:charts 返回 ≥2",
        action="get:charts",
        params={"worksheet_name": "Data"},
        pre_ops=[
            ("insert:chart", {"source_address": "A1:C4", "chart_type": "ColumnClustered", "worksheet_name": "Data"}),
            ("insert:chart", {"source_address": "A1:C4", "chart_type": "Line", "worksheet_name": "Data"}),
        ],
        validator=validate_two_charts,
        tags=["get", "multi"],
    ),
    ExcelCase(
        name="get:charts 空表无图表",
        fixture_name=CHART,
        description="get:charts Blank（无图表）→ 空列表",
        action="get:charts",
        params={"worksheet_name": "Blank"},
        validator=validate_empty,
        tags=["get", "empty"],
    ),
    ExcelCase(
        name="错误码 3010 — 不存在的 worksheet（ELEMENT_NOT_FOUND kind:worksheet）",
        fixture_name=CHART,
        description="get:charts worksheet_name='NoSuch' → 3010 + details.kind=worksheet（oasp#17）",
        action="get:charts",
        params={"worksheet_name": "NoSuchSheet"},
        expect_error_code="3010",
        expect_error_details={"kind": "worksheet"},
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Chart E2E — get:charts", TEST_CASES)
