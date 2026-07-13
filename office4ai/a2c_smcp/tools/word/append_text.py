"""word_append_text MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import TextFormat


class WordAppendTextInput(BaseModel):
    """MCP 输入模型: 追加文本"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    text: str = Field(..., description="Text to append")
    location: Literal["Start", "End"] = Field(
        default="End",
        description="Append location: Start (beginning of document) or End (end of document)",
    )
    format: TextFormat | None = Field(
        None,
        description=(
            "Text formatting. Font attributes go under 'font' (WordFont: bold, italic, underline, size, "
            "name, color, highlightColor); optional 'styleName'. Example: {'font': {'bold': true, 'size': 14}}."
        ),
    )


class WordAppendTextTool(BaseTool):
    """在 Word 文档末尾追加文本"""

    @property
    def name(self) -> str:
        return "word_append_text"

    @property
    def description(self) -> str:
        return (
            "Append text to a Word document at the start or end. "
            "Unlike insert_text which inserts at cursor, this appends to document boundaries. "
            "Default appends to the end of the document."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordAppendTextInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "append:text"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordAppendTextInput
