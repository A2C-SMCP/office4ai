"""excel_get_worksheet_info MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetWorksheetInfoInput(BaseModel):
    """MCP 输入模型: 获取工作表详细信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str | None = Field(
        default=None,
        description="Worksheet name; omit to use the active worksheet",
    )


class ExcelGetWorksheetInfoTool(BaseTool):
    """获取 Excel 工作表的详细信息（已使用范围、表格数、图表数）"""

    @property
    def name(self) -> str:
        return "excel_get_worksheet_info"

    @property
    def description(self) -> str:
        return (
            "Get detailed info about an Excel worksheet: its name, used range (address, row/column count), "
            "table count, and chart count. Omit worksheet_name to inspect the active worksheet. "
            "Use this to understand a sheet's extent and contents before reading or writing ranges."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetWorksheetInfoInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:worksheetInfo"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetWorksheetInfoInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回工作表摘要 | Get tool: return worksheet summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        name = obs.data.get("name", "?")
        used_range = obs.data.get("usedRange", {})
        address = used_range.get("address", "?") if isinstance(used_range, dict) else "?"
        table_count = obs.data.get("tableCount", 0)
        chart_count = obs.data.get("chartCount", 0)
        content = f"Worksheet {name!r}: usedRange={address}, tables={table_count}, charts={chart_count}"
        return {"success": True, "content": content, "data": obs.data}
