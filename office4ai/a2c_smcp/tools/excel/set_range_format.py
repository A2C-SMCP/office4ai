"""excel_set_range_format MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.excel import SetRangeFormatOptions


class ExcelSetRangeFormatInput(BaseModel):
    """MCP 输入模型: 设置范围格式（偏更新）"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address, e.g. 'A1:C1'")
    format: SetRangeFormatOptions = Field(
        ...,
        description=(
            "Format fields to apply (all optional, partial update): font, fill, borders, alignment, "
            "numberFormat. Only the properties you pass are modified. Alignment/underline/border values "
            "are case-sensitive (e.g. 'Center', 'Justify')."
        ),
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelSetRangeFormatTool(BaseTool):
    """设置 Excel 范围的格式（字体/填充/边框/对齐/数字格式，偏更新）"""

    @property
    def name(self) -> str:
        return "excel_set_range_format"

    @property
    def description(self) -> str:
        return (
            "Set the format of an Excel range. All format fields are optional (partial update): only the "
            "properties you pass are modified — font, fill, borders (top/bottom/left/right), alignment "
            "(horizontal/vertical/wrapText/indentLevel/textOrientation) and numberFormat (a single format "
            "string like '0.00' or 'yyyy-mm-dd'). Alignment/underline/border style values are case-sensitive "
            "(e.g. 'Center'). Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelSetRangeFormatInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "set:rangeFormat"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelSetRangeFormatInput
