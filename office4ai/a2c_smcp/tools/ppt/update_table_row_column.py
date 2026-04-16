"""ppt_update_table_row_column MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import ColumnUpdate, RowUpdate


class PptUpdateTableRowColumnInput(BaseModel):
    """MCP 输入模型: 按行/列批量更新表格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | int = Field(..., description="Table element ID")
    rows: list[RowUpdate] | None = Field(None, description="Row-level batch updates")
    columns: list[ColumnUpdate] | None = Field(None, description="Column-level batch updates")

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("elementId", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptUpdateTableRowColumnTool(BaseTool):
    """按行或按列批量更新表格内容"""

    @property
    def name(self) -> str:
        return "ppt_update_table_row_column"

    @property
    def description(self) -> str:
        return (
            "Batch update table content by row or column on a PowerPoint slide. "
            "Provide rows (with rowIndex and values array) and/or columns (with columnIndex and values array). "
            "When both are provided, rows are processed first, then columns may override."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateTableRowColumnInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:tableRowColumn"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateTableRowColumnInput
