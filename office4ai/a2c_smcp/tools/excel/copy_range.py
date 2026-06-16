"""excel_copy_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelCopyRangeInput(BaseModel):
    """MCP 输入模型: 复制范围到目标位置"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    source_address: str = Field(..., description="Source range address, e.g. 'A1:C3'")
    target_address: str = Field(..., description="Target range address, e.g. 'E1:G3'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelCopyRangeTool(BaseTool):
    """复制 Excel 源范围到目标位置"""

    @property
    def name(self) -> str:
        return "excel_copy_range"

    @property
    def description(self) -> str:
        return (
            "Copy an Excel range (values and formatting) from source_address to target_address. "
            "Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelCopyRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "copy:range"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelCopyRangeInput
