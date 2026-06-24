"""excel_add_conditional_format MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.excel import ConditionalFormatRule


class ExcelAddConditionalFormatInput(BaseModel):
    """MCP 输入模型: 添加条件格式规则"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Range address, e.g. 'B2:B100'")
    rule: ConditionalFormatRule = Field(
        ...,
        description=(
            "Conditional-format rule (passthrough). Requires 'type' "
            "(e.g. 'cellValue', 'colorScale', 'dataBar', 'iconSet') plus any rule-specific parameters, "
            "e.g. {'type': 'cellValue', 'operator': 'greaterThan', 'value': 90, "
            "'format': {'fill': {'color': '#C6EFCE'}}}."
        ),
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelAddConditionalFormatTool(BaseTool):
    """为 Excel 范围添加条件格式规则（规则透传）"""

    @property
    def name(self) -> str:
        return "excel_add_conditional_format"

    @property
    def description(self) -> str:
        return (
            "Add a conditional-format rule to an Excel range. The rule is passed through to Excel: it must "
            "include a 'type' (cellValue, colorScale, dataBar, iconSet, ...) and any type-specific "
            "parameters. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelAddConditionalFormatInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "add:conditionalFormat"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelAddConditionalFormatInput
