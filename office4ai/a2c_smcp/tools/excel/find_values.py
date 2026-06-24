"""excel_find_values MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelFindValuesInput(BaseModel):
    """MCP 输入模型: 在指定范围内搜索值"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    search_text: str = Field(..., description="Text to search for")
    address: str | None = Field(
        default=None, description="Search range, e.g. 'A1:D100'; omit to search the whole sheet"
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")
    match_case: bool | None = Field(
        default=None, description="Case-sensitive match; omit for false (the AddIn default)"
    )
    match_entire_cell: bool | None = Field(
        default=None, description="Require the whole cell to match; omit for false (the AddIn default)"
    )


class ExcelFindValuesTool(BaseTool):
    """在 Excel 范围内查找值，返回命中单元格地址与值"""

    @property
    def name(self) -> str:
        return "excel_find_values"

    @property
    def description(self) -> str:
        return (
            "Search for a value within a range and return the address and value of every matching cell. "
            "Omit address to search the whole worksheet. match_case (default false) toggles case sensitivity; "
            "match_entire_cell (default false) requires the whole cell content to match rather than a substring. "
            "Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelFindValuesInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "find:values"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelFindValuesInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回命中摘要 | Get tool: return match summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        matches = obs.data.get("matches", [])
        addresses = [m.get("address", "?") for m in matches if isinstance(m, dict)]
        # 计数基于 addresses（已过滤畸形条目），保证计数与列表始终一致
        content = f"{len(addresses)} match(es): {', '.join(addresses)}" if addresses else "0 match(es)"
        return {"success": True, "content": content, "data": obs.data}
