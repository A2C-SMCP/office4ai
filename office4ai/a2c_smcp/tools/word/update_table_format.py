"""word_update_table_format MCP Tool (OASP /word Draft, v0.2.0)"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import (
    TableBorderOptions,
    TableStyleOptions,
)


class WordUpdateTableFormatInput(BaseModel):
    """MCP 输入模型: 更新 Word 整表格式（样式、边框、列宽、对齐、内边距）"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    table_id: str | None = Field(
        default=None,
        description=(
            "Target table identifier (e.g. 'table-0'). "
            "Omit to target the table containing the current cursor; "
            "if the cursor is not inside a table, the operation fails with 3013 NO_TABLE_AT_CURSOR."
        ),
    )
    style_options: TableStyleOptions | None = Field(
        default=None,
        description=(
            "Style preset (styleType, firstColumn, lastColumn, totalRow, bandedRows, "
            "bandedColumns) plus table-level cellPadding (top/bottom/left/right). "
            "An unknown styleType triggers 3011 STYLE_NOT_FOUND."
        ),
    )
    border_options: TableBorderOptions | None = Field(
        default=None,
        description=(
            "Border options: location ('all' | 'inside' | 'outside', default 'all'), "
            "style ('Single' | 'Double' | 'Dashed' | 'Dotted' | 'None'), "
            "width (points), color (hex)."
        ),
    )
    column_widths: list[float] | None = Field(
        default=None,
        description=(
            "Column widths in points; length must be ≤ table column count (over-length triggers 4002 INVALID_PARAM)."
        ),
    )
    alignment: Literal["Left", "Centered", "Right"] | None = Field(
        default=None,
        description="Whole-table horizontal alignment (Word.Alignment enum: 'Left' | 'Centered' | 'Right').",
    )


class WordUpdateTableFormatTool(BaseTool):
    """Update overall Word table format — preset, borders, column widths, alignment, padding (OASP /word Draft)."""

    @property
    def name(self) -> str:
        return "word_update_table_format"

    @property
    def description(self) -> str:
        return (
            "Update the overall format of a Word table "
            "(OASP /word Draft). "
            "Apply a built-in style preset (styleType + bandedRows/firstColumn/...), "
            "borders (location 'all' | 'inside' | 'outside' + style + width + color), "
            "column widths in points, whole-table alignment, "
            "and table-level cell padding via styleOptions.cellPadding. "
            "Note: this tool does not write cell text — use word_update_table_row_column for that. "
            "tableId is optional — when omitted, the Add-In targets the table containing the cursor; "
            "if the cursor is not inside a table, the call fails with 3013 NO_TABLE_AT_CURSOR. "
            "Use 'Centered' (NOT 'Center') for alignment, per Word.Alignment enum."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordUpdateTableFormatInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "update:tableFormat"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordUpdateTableFormatInput
