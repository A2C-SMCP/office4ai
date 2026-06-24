"""excel_sort_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.excel import SortField


class ExcelSortTableInput(BaseModel):
    """MCP 输入模型: 对表格按指定列多级排序"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    table_id: str = Field(..., description="Table name or ID, e.g. 'Table1'")
    sort_fields: list[SortField] = Field(
        ...,
        description=(
            "Sort keys in priority order. Each key has columnIndex (0-based) and optional ascending "
            "(omit for ascending, the AddIn default)."
        ),
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelSortTableTool(BaseTool):
    """对 Excel 表格按指定列多级排序"""

    @property
    def name(self) -> str:
        return "excel_sort_table"

    @property
    def description(self) -> str:
        return (
            "Sort an Excel Table by one or more columns. sort_fields is a priority-ordered list; each key "
            "has columnIndex (0-based) and an optional ascending flag (omit for ascending). Fails if the "
            "table does not exist or a column index is out of range. Omit worksheet_name to use the active "
            "worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelSortTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "sort:table"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelSortTableInput
