"""excel_clear_conditional_format MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelClearConditionalFormatInput(BaseModel):
    """MCP 输入模型: 清除范围上的所有条件格式"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address, e.g. 'B2:B100'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelClearConditionalFormatTool(BaseTool):
    """清除 Excel 范围上的所有条件格式"""

    @property
    def name(self) -> str:
        return "excel_clear_conditional_format"

    @property
    def description(self) -> str:
        return (
            "Clear all conditional-format rules on an Excel range by address. Omit worksheet_name to use "
            "the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelClearConditionalFormatInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "clear:conditionalFormat"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelClearConditionalFormatInput
