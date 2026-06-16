"""excel_get_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetTableInput(BaseModel):
    """MCP 输入模型: 获取指定表格的详细信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    table_id: str = Field(..., description="Table name or ID, e.g. 'Table1'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelGetTableTool(BaseTool):
    """读取 Excel 单个表格的详细信息（列明细 / 行列数 / 样式）"""

    @property
    def name(self) -> str:
        return "excel_get_table"

    @property
    def description(self) -> str:
        return (
            "Get detailed info for one Excel Table by name or ID: id, address, data row count (excluding "
            "header), column count, the list of columns (name + 0-based index), style name and whether the "
            "header row is shown. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:table"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetTableInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回表格摘要 | Get tool: return table summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        name = obs.data.get("name", "?")
        address = obs.data.get("address", "?")
        row_count = obs.data.get("rowCount", "?")
        column_count = obs.data.get("columnCount", "?")
        content = f"Table '{name}' at {address}: {row_count} row(s) x {column_count} column(s)"
        return {"success": True, "content": content, "data": obs.data}
