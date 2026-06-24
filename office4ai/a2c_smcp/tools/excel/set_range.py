"""excel_set_range MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelSetRangeInput(BaseModel):
    """MCP 输入模型: 设置范围的值"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Target range address, e.g. 'A1:C3'")
    values: Any = Field(
        ...,
        description="A scalar (fills the whole range) or a 2D array matching the range shape",
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelSetRangeTool(BaseTool):
    """设置 Excel 指定范围的值（标量填充或二维数组）"""

    @property
    def name(self) -> str:
        return "excel_set_range"

    @property
    def description(self) -> str:
        return (
            "Write values into an Excel range. Pass a scalar to fill the entire range with one value, "
            "or a 2D array (row-major) matching the range shape. Omit worksheet_name to use the active "
            "worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelSetRangeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "set:range"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelSetRangeInput
