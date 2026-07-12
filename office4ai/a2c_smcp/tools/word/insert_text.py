"""word_insert_text MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import TextFormat


class WordInsertTextInput(BaseModel):
    """MCP 输入模型: 插入文本"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    text: str = Field(..., description="Text to insert")
    location: Literal["Cursor", "Start", "End"] = Field(
        default="Cursor",
        description="Insertion location: Cursor (at current cursor), Start (beginning of document), End (end of document)",
    )
    format: TextFormat | None = Field(
        None,
        description=(
            "Text formatting. Font attributes go under 'font' (WordFont: bold, italic, underline, size, "
            "name, color, highlightColor); optional 'styleName'. Example: {'font': {'bold': true, 'size': 14}}."
        ),
    )


class WordInsertTextTool(BaseTool):
    """在 Word 文档中插入文本"""

    @property
    def name(self) -> str:
        return "word_insert_text"

    @property
    def description(self) -> str:
        return (
            "Insert text into a Word document at the specified location. "
            "Supports formatting options like bold, italic, font size, font name, color, underline, and Word styles. "
            "Default insertion is at the current cursor position."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertTextInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:text"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertTextInput
