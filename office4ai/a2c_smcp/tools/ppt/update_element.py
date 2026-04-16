"""ppt_update_element MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import ElementUpdates


class PptUpdateElementInput(BaseModel):
    """MCP 输入模型: 更新元素位置/尺寸/旋转"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    elementId: str | int = Field(..., description="Element ID to update")
    slideIndex: int | None = Field(None, description="Slide index (0-based), default: current slide", ge=0)
    updates: ElementUpdates = Field(..., description="Geometric updates (left, top, width, height, rotation)")

    # OF4AI-8: LLM 将纯数字字符串 ID（如 "5"）推断为 int，需强转回 str 以通过下游 DTO 校验
    @field_validator("elementId", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptUpdateElementTool(BaseTool):
    """更新元素位置/尺寸/旋转"""

    @property
    def name(self) -> str:
        return "ppt_update_element"

    @property
    def description(self) -> str:
        return (
            "Update an element's position, size, or rotation on a PowerPoint slide. "
            "This is for geometric properties only (left, top, width, height, rotation). "
            "For text content/style changes, use ppt_update_text_box instead."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateElementInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:element"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateElementInput
