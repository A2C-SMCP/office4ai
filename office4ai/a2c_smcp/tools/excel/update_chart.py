"""excel_update_chart MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.excel import ChartUpdateProperties


class ExcelUpdateChartInput(BaseModel):
    """MCP 输入模型: 更新图表属性（偏更新）"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    chart_name: str = Field(..., description="Chart name, e.g. 'Chart 1' (use excel_get_charts to discover)")
    properties: ChartUpdateProperties = Field(
        ...,
        description=(
            "Properties to update (all optional, partial update): title, chartType (case-sensitive type "
            "string), sourceAddress (new data range), position (top/left/width/height in points). Only the "
            "properties you pass are modified."
        ),
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelUpdateChartTool(BaseTool):
    """更新 Excel 图表属性（标题/类型/数据源/位置，偏更新）"""

    @property
    def name(self) -> str:
        return "excel_update_chart"

    @property
    def description(self) -> str:
        return (
            "Update a chart's properties. All fields in properties are optional (partial update): only the "
            "properties you pass are modified — title, chartType (case-sensitive type string), sourceAddress "
            "(new data range) and position (top/left/width/height in points). Fails if the chart does not "
            "exist. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelUpdateChartInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "update:chart"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelUpdateChartInput
