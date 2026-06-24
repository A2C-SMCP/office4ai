"""excel_delete_pivot_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelDeletePivotTableInput(BaseModel):
    """MCP 输入模型: 删除指定透视表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    pivot_table_name: str = Field(..., description="Pivot table name to delete, e.g. 'PivotTable1'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelDeletePivotTableTool(BaseTool):
    """删除 Excel 工作表中的指定透视表"""

    @property
    def name(self) -> str:
        return "excel_delete_pivot_table"

    @property
    def description(self) -> str:
        return (
            "Delete a PivotTable by name from a worksheet. Fails if the pivot table does not exist. Use "
            "excel_get_pivot_tables to discover pivot table names. Omit worksheet_name to use the active "
            "worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelDeletePivotTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "delete:pivotTable"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelDeletePivotTableInput
