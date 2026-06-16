"""excel_get_worksheets MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetWorksheetsInput(BaseModel):
    """MCP 输入模型: 获取工作簿中所有工作表列表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")


class ExcelGetWorksheetsTool(BaseTool):
    """列出 Excel 工作簿中的所有工作表"""

    @property
    def name(self) -> str:
        return "excel_get_worksheets"

    @property
    def description(self) -> str:
        return (
            "List all worksheets in the Excel workbook: each entry has name, 0-based index, and "
            "active/hidden flags. Use this to discover sheet names before deleting, renaming, or "
            "activating a worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetWorksheetsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:worksheets"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetWorksheetsInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回工作表列表摘要 | Get tool: return worksheet-list summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        worksheets = obs.data.get("worksheets", [])
        names = [w.get("name", "?") for w in worksheets if isinstance(w, dict)]
        # 计数基于 names（已过滤畸形条目），保证计数与列表始终一致
        content = f"{len(names)} worksheet(s): {', '.join(names)}" if names else "0 worksheet(s)"
        return {"success": True, "content": content, "data": obs.data}
