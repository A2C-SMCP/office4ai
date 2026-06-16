"""excel_set_formula MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelSetFormulaInput(BaseModel):
    """MCP 输入模型: 设置单元格公式"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Target cell address, e.g. 'D1'")
    formula: str = Field(
        ...,
        description="Formula string including the leading '=' (e.g. '=SUM(A1:C1)', '=VLOOKUP(A1,B:C,2,FALSE)')",
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelSetFormulaTool(BaseTool):
    """设置 Excel 单元格的公式"""

    @property
    def name(self) -> str:
        return "excel_set_formula"

    @property
    def description(self) -> str:
        return (
            "Set the formula of an Excel cell. Pass the full formula string including the leading '=' "
            "(e.g. '=SUM(A1:C1)'). The formula is applied as-is. Omit worksheet_name to use the active "
            "worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelSetFormulaInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "set:formula"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelSetFormulaInput
