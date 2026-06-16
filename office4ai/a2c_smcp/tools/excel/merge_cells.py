"""excel_merge_cells MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelMergeCellsInput(BaseModel):
    """MCP 输入模型: 合并单元格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address to merge, e.g. 'A1:C1'")
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelMergeCellsTool(BaseTool):
    """合并 Excel 单元格（保留左上角单元格的值）"""

    @property
    def name(self) -> str:
        return "excel_merge_cells"

    @property
    def description(self) -> str:
        return (
            "Merge the cells of an Excel range into a single cell. The value of the top-left cell is kept. "
            "Fails with MERGE_CONFLICT if the range conflicts with an existing merged area. Omit "
            "worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelMergeCellsInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "merge:cells"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelMergeCellsInput
