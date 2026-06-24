"""excel_activate_worksheet MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelActivateWorksheetInput(BaseModel):
    """MCP 输入模型: 激活（切换到）指定工作表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    worksheet_name: str = Field(..., description="Name of the worksheet to activate")


class ExcelActivateWorksheetTool(BaseTool):
    """激活（切换到）Excel 指定工作表"""

    @property
    def name(self) -> str:
        return "excel_activate_worksheet"

    @property
    def description(self) -> str:
        return (
            "Activate (switch to) an Excel worksheet by name, making it the active sheet. Fails if the "
            "worksheet does not exist. worksheet_name is required."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelActivateWorksheetInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "activate:worksheet"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelActivateWorksheetInput
