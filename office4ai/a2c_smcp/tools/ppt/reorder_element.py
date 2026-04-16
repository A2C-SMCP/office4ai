"""ppt_reorder_element MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool


class PptReorderElementInput(BaseModel):
    """MCP 输入模型: 调整元素层叠顺序"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | int = Field(..., description="Element ID to reorder")
    slideIndex: int | None = Field(None, description="Slide index (0-based), default: current slide", ge=0)
    action: Literal["bringToFront", "sendToBack", "bringForward", "sendBackward"] = Field(
        ..., description="Reorder action"
    )

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("elementId", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptReorderElementTool(BaseTool):
    """调整元素层叠顺序"""

    @property
    def name(self) -> str:
        return "ppt_reorder_element"

    @property
    def description(self) -> str:
        return (
            "Adjust the z-order (stacking order) of an element on a PowerPoint slide. "
            "Supports bringToFront, sendToBack, bringForward, and sendBackward actions."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptReorderElementInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "reorder:element"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptReorderElementInput
