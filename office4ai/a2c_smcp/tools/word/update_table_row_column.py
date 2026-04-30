"""word_update_table_row_column MCP Tool (OASP /word Draft, v0.2.0)"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import TableColumnUpdate, TableRowUpdate


class WordUpdateTableRowColumnInput(BaseModel):
    """MCP 输入模型: 按行/列批量写入 Word 表格文本"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    table_id: str | None = Field(
        default=None,
        description=(
            "Target table identifier (e.g. 'table-0'). "
            "Omit to target the table containing the current cursor; "
            "if the cursor is not inside a table, the operation fails with 3013 NO_TABLE_AT_CURSOR."
        ),
    )
    rows: list[TableRowUpdate] | None = Field(
        default=None,
        description="Row updates — each row writes a full sequence of column values",
    )
    columns: list[TableColumnUpdate] | None = Field(
        default=None,
        description="Column updates — each column writes a full sequence of row values",
    )

    @model_validator(mode="after")
    def _ensure_rows_or_columns(self) -> WordUpdateTableRowColumnInput:
        if not self.rows and not self.columns:
            raise ValueError("At least one of 'rows' or 'columns' must be provided")
        return self


class WordUpdateTableRowColumnTool(BaseTool):
    """Batch-update Word table cells by entire row(s) or column(s) (OASP /word Draft)."""

    @property
    def name(self) -> str:
        return "word_update_table_row_column"

    @property
    def description(self) -> str:
        return (
            "Batch-write Word table cells by full rows or columns "
            "(OASP /word Draft). "
            "Provide rows[] (each row carries values[] for every column) and/or "
            "columns[] (each column carries values[] for every row). "
            "Typical use: fill several rows of structured contract / report data "
            "in a single round-trip. "
            "tableId is optional — when omitted, the Add-In targets the table containing the cursor; "
            "if the cursor is not inside a table, the call fails with 3013 NO_TABLE_AT_CURSOR. "
            "At least one of rows / columns must be supplied."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordUpdateTableRowColumnInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "update:tableRowColumn"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordUpdateTableRowColumnInput
