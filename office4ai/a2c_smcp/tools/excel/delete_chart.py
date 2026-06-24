"""excel_delete_chart MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelDeleteChartInput(BaseModel):
    """MCP 输入模型: 删除指定图表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    chart_name: str = Field(..., description="Chart name to delete, e.g. 'Chart 1'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelDeleteChartTool(BaseTool):
    """删除 Excel 工作表中的指定图表"""

    @property
    def name(self) -> str:
        return "excel_delete_chart"

    @property
    def description(self) -> str:
        return (
            "Delete a chart by name from a worksheet. Fails if the chart does not exist. Use excel_get_charts "
            "to discover chart names. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelDeleteChartInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "delete:chart"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelDeleteChartInput
