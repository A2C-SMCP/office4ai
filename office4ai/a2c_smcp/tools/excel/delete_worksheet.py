"""excel_delete_worksheet MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelDeleteWorksheetInput(BaseModel):
    """MCP 输入模型: 删除指定工作表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str = Field(..., description="Name of the worksheet to delete")


class ExcelDeleteWorksheetTool(BaseTool):
    """删除 Excel 指定工作表"""

    @property
    def name(self) -> str:
        return "excel_delete_worksheet"

    @property
    def description(self) -> str:
        return (
            "Delete a worksheet from the Excel workbook by name. Fails if the worksheet does not exist "
            "or is protected. worksheet_name is required."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelDeleteWorksheetInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "delete:worksheet"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelDeleteWorksheetInput
