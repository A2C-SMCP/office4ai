"""excel_get_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetRangeInput(BaseModel):
    """MCP 输入模型: 获取范围的值（可选含格式）"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address, e.g. 'A1:C3' or 'Sheet1!A1:C3'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")
    include_format: bool = Field(
        default=False, description="Whether to also return cell format info (font, fill, alignment, ...)"
    )


class ExcelGetRangeTool(BaseTool):
    """读取 Excel 指定范围的值，可选包含格式信息"""

    @property
    def name(self) -> str:
        return "excel_get_range"

    @property
    def description(self) -> str:
        return (
            "Read the values of an Excel range by address (e.g. 'A1:C3' or 'Sheet1!A1:C3'): returns the "
            "2D value array (row-major) plus row/column counts. Set include_format=true to also get "
            "font/fill/alignment/number-format info. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:range"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetRangeInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回范围摘要 | Get tool: return range summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        address = obs.data.get("address", "?")
        row_count = obs.data.get("rowCount", "?")
        column_count = obs.data.get("columnCount", "?")
        has_format = "format" in obs.data
        content = f"Range {address}: {row_count}×{column_count}" + (" (with format)" if has_format else "")
        return {"success": True, "content": content, "data": obs.data}
