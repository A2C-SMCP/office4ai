"""ppt_insert_shape MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.ppt import ShapeInsertOptions


class PptInsertShapeInput(BaseModel):
    """MCP 输入模型: 插入形状"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    shapeType: Literal[
        "Rectangle",
        "RoundedRectangle",
        "Circle",
        "Oval",
        "Triangle",
        "Line",
        "Arrow",
        "Star",
        "TextBox",
    ] = Field(
        ...,
        description=(
            "Shape type. 'TextBox' is a real text box (no fill, no border). "
            "Note: 'Line' is currently unreliable (rendered like a rectangle, office-editor4ai #60) — avoid for now."
        ),
    )
    options: ShapeInsertOptions | None = Field(
        None,
        description=(
            "Shape insertion options (position, size, style, text, and font: PptFont with "
            "size/name/color/bold/italic/underline/...). text/font apply to text-capable shapes "
            "only — passing them on 'Line' (no text box) is rejected with 4002."
        ),
    )


class PptInsertShapeTool(BaseTool):
    """在幻灯片上插入形状"""

    @property
    def name(self) -> str:
        return "ppt_insert_shape"

    @property
    def description(self) -> str:
        return (
            "Insert a geometric shape on a PowerPoint slide. "
            "Supports Rectangle, RoundedRectangle, Circle, Oval, Triangle, "
            "Line, Arrow, Star, and TextBox shape types "
            "(TextBox = a real text box with no fill/border; 'Line' is unreliable, see office-editor4ai #60). "
            "By default shapes have NO fill and NO border; set fillColor / borderColor (hex) to add them, "
            "or pass 'none' / borderWidth=0 to explicitly disable. Supports optional position, size, text, and "
            "font (insert-with-font). text/font are for text-capable shapes only (all except 'Line', which has "
            "no text box); passing text/font on 'Line' is rejected with 4002. When both text and font are given, "
            "font is applied to the inserted text."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptInsertShapeInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "insert:shape"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptInsertShapeInput
