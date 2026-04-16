"""ppt_update_table_format MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import CellFormat, ColumnFormat, RowFormat


class PptUpdateTableFormatInput(BaseModel):
    """MCP 输入模型: 更新表格样式"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | int = Field(..., description="Table element ID")
    cellFormats: list[CellFormat] | None = Field(None, description="Cell-level format settings")
    rowFormats: list[RowFormat] | None = Field(None, description="Row-level format settings")
    columnFormats: list[ColumnFormat] | None = Field(None, description="Column-level format settings")

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("elementId", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptUpdateTableFormatTool(BaseTool):
    """更新表格样式格式"""

    @property
    def name(self) -> str:
        return "ppt_update_table_format"

    @property
    def description(self) -> str:
        return (
            "Update table formatting on a PowerPoint slide. "
            "Supports per-cell, per-row, and per-column formatting including background color, font size, "
            "font color, bold, italic, and alignment. "
            "Priority: cellFormats > columnFormats > rowFormats."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateTableFormatInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:tableFormat"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateTableFormatInput
