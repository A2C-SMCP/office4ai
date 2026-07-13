"""ppt_insert_text MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import TextInsertOptions


class PptInsertTextInput(BaseModel):
    """MCP 输入模型: 插入文本框"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    text: str = Field(..., description="Text to insert")
    options: TextInsertOptions | None = Field(
        None,
        description="Insertion options (position, size, fillColor, font: PptFont with size/name/color/bold/italic/underline/...)",
    )


class PptInsertTextTool(BaseTool):
    """在幻灯片上插入文本框"""

    @property
    def name(self) -> str:
        return "ppt_insert_text"

    @property
    def description(self) -> str:
        return (
            "Insert a text box on a PowerPoint slide. "
            "Supports position, size, font settings, font color, and optional fill/border. "
            "By default the text box has NO fill and NO border (clean look); set fillColor / "
            "borderColor (hex) to add them, or pass 'none' / borderWidth=0 to keep them off. "
            "Default insertion is on the current slide."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptInsertTextInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "insert:text"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptInsertTextInput
