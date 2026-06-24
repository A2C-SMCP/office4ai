"""excel_get_pivot_tables MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetPivotTablesInput(BaseModel):
    """MCP 输入模型: 获取工作表中所有透视表列表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelGetPivotTablesTool(BaseTool):
    """列出 Excel 工作表中所有透视表（name / id）"""

    @property
    def name(self) -> str:
        return "excel_get_pivot_tables"

    @property
    def description(self) -> str:
        return (
            "List all PivotTables in a worksheet: each entry has the pivot table name and id. Use this to "
            "discover pivot table names before deleting one. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetPivotTablesInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:pivotTables"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetPivotTablesInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回透视表列表摘要 | Get tool: return pivot-table-list summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        pivot_tables = obs.data.get("pivotTables", [])
        names = [p.get("name", "?") for p in pivot_tables if isinstance(p, dict)]
        # 计数基于 names（已过滤畸形条目），保证计数与列表始终一致
        content = f"{len(names)} pivot table(s): {', '.join(names)}" if names else "0 pivot table(s)"
        return {"success": True, "content": content, "data": obs.data}
