"""excel_insert_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelInsertRangeInput(BaseModel):
    """MCP 输入模型: 插入空白单元格并移动现有单元格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address to insert blank cells at, e.g. 'B2:B5'")
    shift_direction: Literal["down", "right"] = Field(
        ...,
        description="Direction existing cells shift to make room: 'down' or 'right'",
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelInsertRangeTool(BaseTool):
    """在 Excel 指定位置插入空白单元格，将现有单元格向下或向右移动"""

    @property
    def name(self) -> str:
        return "excel_insert_range"

    @property
    def description(self) -> str:
        return (
            "Insert blank cells into an Excel range, shifting existing cells to make room. "
            "shift_direction is 'down' (existing cells move down) or 'right' (existing cells move right). "
            "Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelInsertRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "insert:range"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelInsertRangeInput
