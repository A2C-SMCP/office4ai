"""excel_add_worksheet MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelAddWorksheetInput(BaseModel):
    """MCP 输入模型: 添加新工作表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    name: str | None = Field(default=None, description="New worksheet name; omit to let Excel auto-name it")


class ExcelAddWorksheetTool(BaseTool):
    """向 Excel 工作簿添加新工作表"""

    @property
    def name(self) -> str:
        return "excel_add_worksheet"

    @property
    def description(self) -> str:
        return (
            "Add a new worksheet to the Excel workbook. Provide name to set its title, or omit to let "
            "Excel auto-name it. Returns the actual name and index of the new worksheet."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelAddWorksheetInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "add:worksheet"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelAddWorksheetInput
