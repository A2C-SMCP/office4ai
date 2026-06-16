"""excel_insert_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelInsertTableInput(BaseModel):
    """MCP 输入模型: 在指定范围创建结构化表格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Table range address, e.g. 'A1:C4'")
    has_headers: bool = Field(..., description="Whether the first row is a header row")
    data: list[list[Any]] | None = Field(
        default=None, description="Initial data as a 2D array (rows of cells); omit to create an empty table"
    )
    style_name: str | None = Field(
        default=None, description="Table style name, e.g. 'TableStyleMedium2'; omit for the default style"
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelInsertTableTool(BaseTool):
    """在 Excel 指定范围创建结构化表格（Excel Table）"""

    @property
    def name(self) -> str:
        return "excel_insert_table"

    @property
    def description(self) -> str:
        return (
            "Create a structured Excel Table over a range. Set has_headers to mark the first row as a "
            "header row. Optionally provide data (a 2D array) to fill the table and style_name to apply a "
            "built-in table style (e.g. 'TableStyleMedium2'). Returns the auto-generated table name and its "
            "actual range. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelInsertTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "insert:table"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelInsertTableInput
