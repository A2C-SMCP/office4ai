"""excel_insert_chart MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.excel import COMMON_CHART_TYPES, ChartPosition


class ExcelInsertChartInput(BaseModel):
    """MCP 输入模型: 根据数据范围创建图表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    source_address: str = Field(..., description="Data source range, e.g. 'A1:C4'")
    chart_type: str = Field(
        ...,
        description=f"Chart type (case-sensitive open string; common values: {COMMON_CHART_TYPES})",
    )
    title: str | None = Field(default=None, description="Chart title; omit for no title")
    position: ChartPosition | None = Field(
        default=None,
        description=(
            "Position and size in points: top, left, width, height (all optional); omit to let Excel "
            "auto-place the chart."
        ),
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelInsertChartTool(BaseTool):
    """在 Excel 中根据数据范围创建图表"""

    @property
    def name(self) -> str:
        return "excel_insert_chart"

    @property
    def description(self) -> str:
        return (
            "Create a chart from a data range. source_address is the data range (e.g. 'A1:C4') and "
            "chart_type is a case-sensitive type string (see the chart_type parameter for the common "
            "values, e.g. ColumnClustered, Line, Pie, XYScatter). Optionally set title and position "
            "(top/left/width/height in points). Returns the auto-generated chart name. Omit worksheet_name "
            "to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelInsertChartInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "insert:chart"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelInsertChartInput
