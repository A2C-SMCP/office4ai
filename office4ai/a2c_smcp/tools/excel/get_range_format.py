"""excel_get_range_format MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetRangeFormatInput(BaseModel):
    """MCP 输入模型: 获取范围完整格式信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address, e.g. 'A1:C3' or 'Sheet1!A1:C3'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelGetRangeFormatTool(BaseTool):
    """读取 Excel 范围的完整格式信息（字体/填充/对齐/数字格式）"""

    @property
    def name(self) -> str:
        return "excel_get_range_format"

    @property
    def description(self) -> str:
        return (
            "Read the full format info of an Excel range by address: font (name/size/bold/italic/color/"
            "underline), fill color, horizontal/vertical alignment, wrap-text and per-cell number formats. "
            "Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetRangeFormatInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:rangeFormat"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetRangeFormatInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回格式摘要 | Get tool: return format summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        address = obs.data.get("address", "?")
        fmt = obs.data.get("format", {})
        font = fmt.get("font", {}) if isinstance(fmt, dict) else {}
        font_desc = font.get("name", "?")
        if "size" in font:
            font_desc += f" {font['size']}pt"
        content = f"Format of {address}: font {font_desc}"
        return {"success": True, "content": content, "data": obs.data}
