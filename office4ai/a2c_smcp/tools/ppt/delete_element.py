"""ppt_delete_element MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool


class PptDeleteElementInput(BaseModel):
    """MCP 输入模型: 删除元素"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | int | None = Field(None, description="Single element ID to delete")
    elementIds: list[str | int] | None = Field(None, description="Batch element IDs to delete")
    slideIndex: int | None = Field(None, description="Slide index (0-based), default: current slide", ge=0)

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("elementId", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v

    @field_validator("elementIds", mode="before")
    @classmethod
    def _coerce_element_ids(cls, v: Any) -> Any:
        if isinstance(v, list):
            return [str(x) if isinstance(x, int) else x for x in v]
        return v


class PptDeleteElementTool(BaseTool):
    """删除幻灯片上的元素"""

    @property
    def name(self) -> str:
        return "ppt_delete_element"

    @property
    def description(self) -> str:
        return (
            "Delete one or more elements from a PowerPoint slide. "
            "Provide elementId for single deletion or elementIds for batch deletion. "
            "If both are provided, elementIds takes priority."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptDeleteElementInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "delete:element"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptDeleteElementInput
