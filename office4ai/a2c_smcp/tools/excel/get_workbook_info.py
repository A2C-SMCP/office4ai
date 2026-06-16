"""excel_get_workbook_info MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs


class ExcelGetWorkbookInfoInput(BaseModel):
    """MCP 输入模型: 获取工作簿信息"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")


class ExcelGetWorkbookInfoTool(BaseTool):
    """获取 Excel 工作簿信息（工作表列表、活动工作表、文件名）"""

    @property
    def name(self) -> str:
        return "excel_get_workbook_info"

    @property
    def description(self) -> str:
        return (
            "Get the current Excel workbook's basic info: the list of all worksheets "
            "(name, 0-based index, active/hidden flags), the active worksheet name, and the file name. "
            "Use this first to orient yourself before operating on a specific sheet or range."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelGetWorkbookInfoInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "get:workbookInfo"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelGetWorkbookInfoInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回工作簿摘要 | Get tool: return workbook summary"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        sheets = obs.data.get("sheets", [])
        active = obs.data.get("activeSheet", "?")
        file_name = obs.data.get("fileName", "?")
        content = f"Workbook {file_name!r}: {len(sheets)} sheet(s), active={active!r}"
        return {"success": True, "content": content, "data": obs.data}
