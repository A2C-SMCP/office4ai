"""excel_unmerge_cells MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelUnmergeCellsInput(BaseModel):
    """MCP 输入模型: 取消单元格合并"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address to unmerge, e.g. 'A1:C1'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelUnmergeCellsTool(BaseTool):
    """取消 Excel 单元格合并"""

    @property
    def name(self) -> str:
        return "excel_unmerge_cells"

    @property
    def description(self) -> str:
        return "Unmerge the merged cells in an Excel range by address. Omit worksheet_name to use the active worksheet."

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelUnmergeCellsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "unmerge:cells"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelUnmergeCellsInput
