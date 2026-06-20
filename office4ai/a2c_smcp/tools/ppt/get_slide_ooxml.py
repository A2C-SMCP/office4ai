"""ppt_get_slide_ooxml MCP Tool (OASP /ppt, v0.3.0 — whole-slide OOXML).

Export one slide's live OOXML as a ``.pptx`` package. Companion to
``ppt_insert_slides_ooxml``: the returned opaque ``slideId`` feeds the latter's
``replaceSlideId`` for in-place round-trip editing.

Context-cost design (office4ai #40)
===================================
The exported base64 is written to ``dest_path`` on disk; the tool returns only
``{slideId, slideIndex, filePath}`` — the (large) base64 never enters the model's
conversation context.
"""

from __future__ import annotations

import base64 as base64lib
from pathlib import Path
from typing import Any, Literal

from loguru import logger
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeAction


class PptGetSlideOoxmlInput(BaseModel):
    """MCP 输入模型: 导出整页 OOXML 到文件 (不回 inline base64)。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    slide_index: int = Field(..., alias="slideIndex", ge=0, description="Target slide index (0-based)")
    dest_path: str = Field(
        ...,
        description=(
            "Filesystem path to write the exported single-slide .pptx package. "
            "The base64 is decoded to disk here and is NOT returned inline."
        ),
    )

    model_config = {"populate_by_name": True}


class PptGetSlideOoxmlTool(BaseTool):
    """导出指定页当前 OOXML 为 .pptx 文件，返回不透明 slideId 供后续替换定位。"""

    @property
    def name(self) -> str:
        return "ppt_get_slide_ooxml"

    @property
    def description(self) -> str:
        return (
            "Export one slide's live OOXML (including unsaved local edits) as a single-slide .pptx "
            "package written to `dest_path` on disk. Returns {slideId, slideIndex, filePath} — the "
            "base64 is NOT returned inline (it stays on disk to keep it out of the model context). "
            "The opaque `slideId` can be passed as `replaceSlideId` to ppt_insert_slides_ooxml for "
            "in-place round-trip editing. Requires the document open in PowerPoint via the Add-In "
            "(PowerPointApi 1.8)."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetSlideOoxmlInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:slideOoxml"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetSlideOoxmlInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Emit ppt:get:slideOoxml → decode the returned base64 to ``dest_path`` → return the handle."""
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, PptGetSlideOoxmlInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 触发导出 (走通用 execute 路径: 连接检查 + emit + 解包)
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={"document_uri": validated.document_uri, "slideIndex": validated.slide_index},
        )
        try:
            obs = await self.workspace.execute(action)
        except Exception as e:  # noqa: BLE001
            logger.exception(f"工具执行失败 | Tool execution failed: {self.name}")
            return {"success": False, "error": f"3004: {e}"}

        if not obs.success:
            return self.format_result(obs)

        # 3. 把导出的 base64 落盘 (不回传 inline base64)
        wire_base64 = obs.data.get("base64") if isinstance(obs.data, dict) else None
        if not wire_base64:
            return {"success": False, "error": "3004: export succeeded but no base64 in response"}
        dest = Path(validated.dest_path)
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(base64lib.b64decode(wire_base64))
        except (OSError, ValueError) as e:
            return {"success": False, "error": f"3004: failed to write dest_path: {e}"}

        # 4. 活动追踪 + 返回句柄 (不含 base64)
        self.workspace.update_last_activity(
            document_uri=validated.document_uri,
            tool_name=self.name,
            result_data={"slideId": obs.data.get("slideId"), "slideIndex": obs.data.get("slideIndex")},
        )
        return {
            "success": True,
            "data": {
                "slideId": obs.data.get("slideId"),
                "slideIndex": obs.data.get("slideIndex"),
                "filePath": str(dest),
            },
        }
