"""
Excel Chart E2E — insert:chart（创建图表）

覆盖 ``excel:insert:chart`` 的多种 chartType（ColumnClustered/Line/Pie/XYScatter）+ title。

wire 形态（AddIn chart.ts 确认）:
- insert:chart（sourceAddress, chartType, title?, position?, worksheetName?）→ ``{name}``
  （name 由 Excel 自动命名，中文环境如 "图表 1"）。

图表为视觉对象：主验证为协议返回非空 ``name``；openpyxl ``chart_count`` 仅 best-effort
（openpyxl 对 Excel 原生图表回读支持有限，读不到不判失败）。富属性回读见 ``test_get_charts.py``。
含错误码：非法 chartType（非空但 Office.js 拒绝枚举）→ 3000 DOCUMENT_ERROR。

运行方式:
    uv run python manual_tests/excel/chart_e2e/test_insert_chart.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.chart_e2e._fixtures import ensure_fixtures
from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main

CHART = "chart_e2e/chart.xlsx"


def _make_validator(label: str):
    def _v(data: dict[str, Any], reader: WorkbookReader) -> bool:
        name = data.get("name")
        print(f"   insert 返回: name={name!r}")
        if not name:
            print("   ❌ 响应缺少非空 name")
            return False
        reader.reload()
        try:
            n = reader.chart_count("Data")
            print(f"   openpyxl chart_count(Data)={n}（best-effort）")
            if n < 1:
                print("   ⚠️  openpyxl 未读到图表（原生图表回读受限；协议已返回 name）")
        except Exception as exc:  # noqa: BLE001
            print(f"   ⚠️  chart_count 读取失败（{exc}）；协议已返回 name")
        print(f"   ✅ {label} 图表已创建（name={name!r}）")
        return True

    return _v


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="insert ColumnClustered + title",
        fixture_name=CHART,
        description="insert:chart Data!A1:C4 ColumnClustered title='Sales vs Cost' → {name}",
        action="insert:chart",
        params={
            "source_address": "A1:C4",
            "chart_type": "ColumnClustered",
            "title": "Sales vs Cost",
            "worksheet_name": "Data",
        },
        validator=_make_validator("ColumnClustered"),
        tags=["insert", "column"],
    ),
    ExcelCase(
        name="insert Line",
        fixture_name=CHART,
        description="insert:chart Data!A1:C4 Line → {name}",
        action="insert:chart",
        params={"source_address": "A1:C4", "chart_type": "Line", "worksheet_name": "Data"},
        validator=_make_validator("Line"),
        tags=["insert", "line"],
    ),
    ExcelCase(
        name="insert Pie（单列）",
        fixture_name=CHART,
        description="insert:chart Data!A1:B4 Pie → {name}",
        action="insert:chart",
        params={"source_address": "A1:B4", "chart_type": "Pie", "worksheet_name": "Data"},
        validator=_make_validator("Pie"),
        tags=["insert", "pie"],
    ),
    ExcelCase(
        name="insert XYScatter + position",
        fixture_name=CHART,
        description="insert:chart Data!A1:C4 XYScatter position{top,left} → {name}",
        action="insert:chart",
        params={
            "source_address": "A1:C4",
            "chart_type": "XYScatter",
            "position": {"top": 200, "left": 300, "width": 360, "height": 240},
            "worksheet_name": "Data",
        },
        validator=_make_validator("XYScatter"),
        tags=["insert", "scatter"],
    ),
    ExcelCase(
        name="错误码 3000 — 非法 chartType（DOCUMENT_ERROR）",
        fixture_name=CHART,
        description="insert:chart chartType='NotARealType'（非空但非法枚举）→ 3000",
        action="insert:chart",
        params={"source_address": "A1:C4", "chart_type": "NotARealType", "worksheet_name": "Data"},
        expect_error_code="3000",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Chart E2E — insert:chart", TEST_CASES)
