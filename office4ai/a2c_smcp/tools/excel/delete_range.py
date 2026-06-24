"""excel_delete_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelDeleteRangeInput(BaseModel):
    """MCP 输入模型: 删除范围并移动周围单元格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address to delete, e.g. 'B2:B5'")
    shift_direction: Literal["up", "left"] = Field(
        ...,
        description="Direction surrounding cells shift to fill the gap: 'up' or 'left'",
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelDeleteRangeTool(BaseTool):
    """删除 Excel 范围并将周围单元格向上或向左移动"""

    @property
    def name(self) -> str:
        return "excel_delete_range"

    @property
    def description(self) -> str:
        return (
            "Delete an Excel range and shift surrounding cells to fill the gap. shift_direction is "
            "'up' (cells below move up) or 'left' (cells to the right move left). Omit worksheet_name "
            "to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelDeleteRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "delete:range"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelDeleteRangeInput
