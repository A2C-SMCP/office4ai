"""excel_get_tables MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetTablesInput(BaseModel):
    """MCP 输入模型: 获取工作表中所有表格的概要列表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelGetTablesTool(BaseTool):
    """列出 Excel 工作表中所有表格的概要（name / id / range）"""

    @property
    def name(self) -> str:
        return "excel_get_tables"

    @property
    def description(self) -> str:
        return (
            "List all structured tables in a worksheet: each entry has the table name, id and range "
            "address. Use this to discover table names before getting details, adding/deleting rows, or "
            "sorting. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetTablesInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:tables"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetTablesInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回表格列表摘要 | Get tool: return table-list summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        tables = obs.data.get("tables", [])
        names = [t.get("name", "?") for t in tables if isinstance(t, dict)]
        # 计数基于 names（已过滤畸形条目），保证计数与列表始终一致
        content = f"{len(names)} table(s): {', '.join(names)}" if names else "0 table(s)"
        return {"success": True, "content": content, "data": obs.data}
