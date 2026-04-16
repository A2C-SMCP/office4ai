"""ppt_update_table_cell MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import TableCellUpdate


class PptUpdateTableCellInput(BaseModel):
    """MCP 输入模型: 更新表格单元格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | int = Field(..., description="Table element ID")
    cells: list[TableCellUpdate] = Field(..., description="Cells to update", min_length=1)

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("elementId", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptUpdateTableCellTool(BaseTool):
    """更新表格中指定单元格的文本内容"""

    @property
    def name(self) -> str:
        return "ppt_update_table_cell"

    @property
    def description(self) -> str:
        return (
            "Update specific cells in a table on a PowerPoint slide. "
            "Each cell update specifies rowIndex, columnIndex, and new text. "
            "Supports batch updates of multiple cells in a single call."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateTableCellInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:tableCell"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateTableCellInput
