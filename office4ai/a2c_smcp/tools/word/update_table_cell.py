"""word_update_table_cell MCP Tool (OASP /word Draft, v0.2.0)"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import TableCellUpdate


class WordUpdateTableCellInput(BaseModel):
    """MCP 输入模型: 更新 Word 表格中指定单元格的文本和/或格式"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    table_id: str | None = Field(
        default=None,
        description=(
            "Target table identifier (e.g. 'table-0'). "
            "Omit to target the table containing the current cursor; "
            "if the cursor is not inside a table, the operation fails with 3013 NO_TABLE_AT_CURSOR."
        ),
    )
    cells: list[TableCellUpdate] = Field(
        ...,
        description=(
            "Cells to update. Each entry needs rowIndex / columnIndex and at least one of "
            "text / format. format = { backgroundColor, font (WordFont: bold, italic, underline, size, "
            "name, color, highlightColor), horizontalAlignment ('Left' | 'Centered' | 'Right' | 'Justified'), "
            "verticalAlignment ('Top' | 'Center' | 'Bottom') }."
        ),
        min_length=1,
    )


class WordUpdateTableCellTool(BaseTool):
    """Update specific cells in a Word table (OASP /word Draft)."""

    @property
    def name(self) -> str:
        return "word_update_table_cell"

    @property
    def description(self) -> str:
        return (
            "Update specific cells in a Word table — text and/or format "
            "(OASP /word Draft). "
            "Each cell is addressed by (rowIndex, columnIndex) and may set new text, "
            "background color, font, alignment, etc. "
            "Use this to apply per-cell visual effects such as a blue-bg centered bold "
            "header cell or alternating gray label / white input cells. "
            "tableId is optional — when omitted, the Add-In targets the table containing the cursor; "
            "if the cursor is not inside a table, the call fails with 3013 NO_TABLE_AT_CURSOR. "
            "Use 'Centered' / 'Justified' (NOT 'Center' / 'Justify') for horizontalAlignment, "
            "and 'Center' (NOT 'Middle') for verticalAlignment, per Word.Alignment enum."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordUpdateTableCellInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "update:tableCell"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordUpdateTableCellInput
