"""excel_set_auto_filter MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.excel import AutoFilterCriterion


class ExcelSetAutoFilterInput(BaseModel):
    """MCP 输入模型: 对指定范围应用自动筛选"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    address: str = Field(..., description="Filter range address, e.g. 'A1:C10'")
    criteria: list[AutoFilterCriterion] = Field(
        ...,
        description=(
            "Per-column filter criteria. Each has columnIndex (0-based), filterOn (filter mode, e.g. 'Values', "
            "'CellColor', 'FontColor'), and optional values (the list of values to keep when filterOn is 'Values')."
        ),
    )
    worksheet_name: str | None = Field(default=None, description="Worksheet name; omit to use the active worksheet")


class ExcelSetAutoFilterTool(BaseTool):
    """对 Excel 范围应用自动筛选"""

    @property
    def name(self) -> str:
        return "excel_set_auto_filter"

    @property
    def description(self) -> str:
        return (
            "Apply an auto-filter to a range. criteria is a per-column list; each entry has columnIndex (0-based), "
            "filterOn (the filter mode, e.g. 'Values', 'CellColor', 'FontColor'), and optional values (the values "
            "to keep when filterOn is 'Values'). Fails if the worksheet is missing, the range is invalid, or a "
            "column index is out of range. Omit worksheet_name to use the active worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelSetAutoFilterInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "set:autoFilter"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelSetAutoFilterInput
