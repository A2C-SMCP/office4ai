"""excel_clear_auto_filter MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelClearAutoFilterInput(BaseModel):
    """MCP 输入模型: 清除工作表上的自动筛选"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelClearAutoFilterTool(BaseTool):
    """清除 Excel 工作表上的自动筛选"""

    @property
    def name(self) -> str:
        return "excel_clear_auto_filter"

    @property
    def description(self) -> str:
        return (
            "Clear the auto-filter on a worksheet, removing all column filters. Omit worksheet_name to use the "
            "active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelClearAutoFilterInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "clear:autoFilter"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelClearAutoFilterInput
