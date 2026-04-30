"""word_merge_cells MCP Tool (OASP /word Draft, v0.2.0)"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class WordMergeCellsInput(BaseModel):
    """MCP 输入模型: 合并 Word 表格中的矩形单元格区域"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    table_id: str | None = Field(
        default=None,
        description=(
            "Target table identifier (e.g. 'table-0'). "
            "Omit to target the table containing the current cursor; "
            "if the cursor is not inside a table, the operation fails with 3013 NO_TABLE_AT_CURSOR."
        ),
    )
    start_row_index: int = Field(
        ...,
        description="Start row index (0-based, inclusive)",
        ge=0,
    )
    start_column_index: int = Field(
        ...,
        description="Start column index (0-based, inclusive)",
        ge=0,
    )
    end_row_index: int = Field(
        ...,
        description="End row index (0-based, inclusive)",
        ge=0,
    )
    end_column_index: int = Field(
        ...,
        description="End column index (0-based, inclusive)",
        ge=0,
    )


class WordMergeCellsTool(BaseTool):
    """Merge a rectangular range of cells in a Word table (OASP /word Draft)."""

    @property
    def name(self) -> str:
        return "word_merge_cells"

    @property
    def description(self) -> str:
        return (
            "Merge a rectangular range of cells in a Word table into a single cell "
            "(OASP /word Draft). "
            "Specify the inclusive (startRowIndex, startColumnIndex) and (endRowIndex, endColumnIndex) "
            "to delimit the region. "
            "Typical use: turn the first row of a table into a single full-width header cell "
            "(e.g. for contract / report headings such as 'Party A Information'). "
            "tableId is optional — when omitted, the Add-In targets the table containing the cursor; "
            "if the cursor is not inside a table, the call fails with 3013 NO_TABLE_AT_CURSOR. "
            "Conflicts with existing merged regions raise 3014 ALREADY_MERGED."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordMergeCellsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "merge:cells"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordMergeCellsInput
