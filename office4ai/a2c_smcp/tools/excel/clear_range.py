"""excel_clear_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelClearRangeInput(BaseModel):
    """MCP 输入模型: 清除范围内容/格式/全部"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address to clear, e.g. 'A1:C3'")
    clear_type: Literal["contents", "formats", "all"] = Field(
        ...,
        description="What to clear: 'contents' (values only), 'formats' (formatting only), or 'all'",
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelClearRangeTool(BaseTool):
    """清除 Excel 范围的内容、格式或全部"""

    @property
    def name(self) -> str:
        return "excel_clear_range"

    @property
    def description(self) -> str:
        return (
            "Clear an Excel range. clear_type selects what to remove: 'contents' clears values only, "
            "'formats' clears formatting only, 'all' clears both. Omit worksheet_name to use the active "
            "worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelClearRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "clear:range"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelClearRangeInput
