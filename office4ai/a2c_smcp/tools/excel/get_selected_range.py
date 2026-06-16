"""excel_get_selected_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetSelectedRangeInput(BaseModel):
    """MCP 输入模型: 获取当前选中范围"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")


class ExcelGetSelectedRangeTool(BaseTool):
    """获取 Excel 当前选中的单元格范围及其值"""

    @property
    def name(self) -> str:
        return "excel_get_selected_range"

    @property
    def description(self) -> str:
        return (
            "Get the currently selected cell range in Excel: its address, the 2D array of cell values "
            "(row-major), and the row/column counts. Use this to read what the user has highlighted."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetSelectedRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:selectedRange"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetSelectedRangeInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回选中范围摘要 | Get tool: return selected-range summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        address = obs.data.get("address", "?")
        row_count = obs.data.get("rowCount", "?")
        column_count = obs.data.get("columnCount", "?")
        content = f"Selected range {address}: {row_count}×{column_count}"
        return {"success": True, "content": content, "data": obs.data}
