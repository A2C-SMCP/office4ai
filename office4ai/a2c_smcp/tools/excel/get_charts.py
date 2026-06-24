"""excel_get_charts MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetChartsInput(BaseModel):
    """MCP 输入模型: 获取工作表中所有图表的信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelGetChartsTool(BaseTool):
    """列出 Excel 工作表中所有图表（name / chartType / title / 位置尺寸）"""

    @property
    def name(self) -> str:
        return "excel_get_charts"

    @property
    def description(self) -> str:
        return (
            "List all charts in a worksheet: each entry has the chart name, chartType, title and position "
            "(top/left/width/height in points). Use this to discover chart names before updating or "
            "deleting a chart. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetChartsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:charts"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetChartsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回图表列表摘要 | Get tool: return chart-list summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        charts = obs.data.get("charts", [])
        names = [c.get("name", "?") for c in charts if isinstance(c, dict)]
        # 计数基于 names（已过滤畸形条目），保证计数与列表始终一致
        content = f"{len(names)} chart(s): {', '.join(names)}" if names else "0 chart(s)"
        return {"success": True, "content": content, "data": obs.data}
