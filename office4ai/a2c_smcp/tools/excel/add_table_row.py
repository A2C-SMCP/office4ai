"""excel_add_table_row MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelAddTableRowInput(BaseModel):
    """MCP 输入模型: 向表格末尾追加一行数据"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    table_id: str = Field(..., description="Table name or ID, e.g. 'Table1'")
    values: list[Any] = Field(..., description="Row values as a 1D array, in column order")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelAddTableRowTool(BaseTool):
    """向 Excel 表格末尾追加一行数据"""

    @property
    def name(self) -> str:
        return "excel_add_table_row"

    @property
    def description(self) -> str:
        return (
            "Append a single row to the end of an Excel Table. values is a 1D array whose elements map to "
            "the table columns in order. Fails if the table does not exist or a value type mismatches the "
            "column. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelAddTableRowInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "add:tableRow"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelAddTableRowInput
