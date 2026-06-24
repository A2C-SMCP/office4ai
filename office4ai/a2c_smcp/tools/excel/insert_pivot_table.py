"""excel_insert_pivot_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelInsertPivotTableInput(BaseModel):
    """MCP 输入模型: 基于数据源范围创建透视表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    source_address: str = Field(..., description="Data source range, e.g. 'A1:D100'")
    target_address: str = Field(..., description="Top-left cell to place the pivot table, e.g. 'F1'")
    name: str | None = Field(default=None, description="Pivot table name; omit for an auto-generated name")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelInsertPivotTableTool(BaseTool):
    """在 Excel 中基于数据源范围创建透视表"""

    @property
    def name(self) -> str:
        return "excel_insert_pivot_table"

    @property
    def description(self) -> str:
        return (
            "Create a PivotTable from a data range. source_address is the data range (e.g. 'A1:D100') and "
            "target_address is the top-left cell where the pivot table is placed (e.g. 'F1'). The pivot table "
            "is created from the source range; this operation does not set the row/column/value field layout. "
            "Optionally set name; omit it for an auto-generated name. Returns the pivot table name. Omit "
            "worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelInsertPivotTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "insert:pivotTable"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelInsertPivotTableInput
