"""excel_rename_worksheet MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool


class ExcelRenameWorksheetInput(BaseModel):
    """MCP 输入模型: 重命名工作表"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/data.xlsx)")
    current_name: str = Field(..., description="Current worksheet name")
    new_name: str = Field(..., description="New worksheet name")


class ExcelRenameWorksheetTool(BaseTool):
    """重命名 Excel 工作表"""

    @property
    def name(self) -> str:
        return "excel_rename_worksheet"

    @property
    def description(self) -> str:
        return (
            "Rename an Excel worksheet from current_name to new_name. Fails if current_name does not "
            "exist or new_name is already taken. Both current_name and new_name are required."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return ExcelRenameWorksheetInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "excel"

    @property
    def event_name(self) -> str:
        return "rename:worksheet"

    @property
    def input_model(self) -> type[BaseModel]:
        return ExcelRenameWorksheetInput
