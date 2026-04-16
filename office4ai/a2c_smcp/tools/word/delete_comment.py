"""word_delete_comment MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool


class WordDeleteCommentInput(BaseModel):
    """MCP 输入模型: 删除批注"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    comment_id: str | int = Field(..., description="ID of the comment to delete")

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("comment_id", mode="before")
    @classmethod
    def _coerce_comment_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class WordDeleteCommentTool(BaseTool):
    """删除 Word 文档中的批注"""

    @property
    def name(self) -> str:
        return "word_delete_comment"

    @property
    def description(self) -> str:
        return (
            "Delete a comment (annotation) from a Word document. "
            "Requires the comment ID, which can be obtained from word_get_comments."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordDeleteCommentInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "delete:comment"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordDeleteCommentInput
