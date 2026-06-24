"""excel_delete_table_row MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelDeleteTableRowInput(BaseModel):
    """MCP 输入模型: 删除表格中指定索引的行"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    table_id: str = Field(..., description="Table name or ID, e.g. 'Table1'")
    row_index: int = Field(..., description="Row index to delete (0-based, excluding the header row)")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelDeleteTableRowTool(BaseTool):
    """删除 Excel 表格中指定索引的行"""

    @property
    def name(self) -> str:
        return "excel_delete_table_row"

    @property
    def description(self) -> str:
        return (
            "Delete one row from an Excel Table by 0-based row index (excluding the header row). Fails if "
            "the table does not exist or the index is out of range. Omit worksheet_name to use the active "
            "worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelDeleteTableRowInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "delete:tableRow"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelDeleteTableRowInput
