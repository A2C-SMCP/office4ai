"""ppt_insert_slides_ooxml MCP Tool (OASP /ppt, v0.3.0 — whole-slide OOXML).

High-fidelity whole-slide insertion: take a ``.pptx`` package built offline
(e.g. via ``python-pptx``) and apply it to the active presentation as REAL native
objects — bypassing the Office.js styling ceiling (no Chart class, no line/paragraph
spacing, no image rounding/crop/shadow).

Context-cost design (office4ai #40)
===================================
The model passes a **filesystem path** (``source_path``), NOT inline base64. The
Server reads the file's bytes and base64-encodes them itself before emitting
``ppt:insert:slidesOoxml`` over the wire to the Add-In. A whole-slide ``.pptx`` is
tens of KB → ~8-28k tokens of base64; keeping it on disk (Server ↔ file ↔ Add-In)
means it never enters the model's conversation context. The base64 on the wire is a
Server→Add-In concern, unrelated to the model context.
"""

from __future__ import annotations

import base64 as base64lib
from pathlib import Path
from typing import Any, Literal

from loguru import logger
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.services.document_lock import document_lock_manager

#: camelCase on the wire; the Add-In maps to PowerPoint.InsertSlideFormatting internally.
SlideFormatting = Literal["keepSourceFormatting", "useDestinationTheme"]


class PptInsertSlidesOoxmlInput(BaseModel):
    """MCP 输入模型: 整页 OOXML 插入 (文件句柄, 非 inline base64)。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    source_path: str = Field(
        ...,
        description=(
            "Filesystem path to a .pptx package (>=1 slide) produced offline (e.g. python-pptx). "
            "The Server reads its bytes and base64-encodes them — DO NOT pass base64 inline. "
            "Use this for high-fidelity whole-slide layouts that the fine-grained tools cannot "
            "achieve; for small per-element tweaks prefer ppt_insert_text / ppt_update_element."
        ),
    )
    target_slide_index: int | None = Field(
        default=None,
        alias="targetSlideIndex",
        ge=0,
        description="Insert after this 0-based index; default = end of document",
    )
    formatting: SlideFormatting | None = Field(
        default=None,
        description="keepSourceFormatting (default) keeps the source look; useDestinationTheme re-themes",
    )
    replace_slide_id: str | None = Field(
        default=None,
        alias="replaceSlideId",
        description="Opaque slideId (from ppt_get_slide_ooxml) to delete after insert (in-place replace)",
    )
    final_slide_index: int | None = Field(
        default=None,
        alias="finalSlideIndex",
        ge=0,
        description="Move the inserted slide to this 0-based index (reposition after replace)",
    )

    model_config = {"populate_by_name": True}


class PptInsertSlidesOoxmlTool(BaseTool):
    """在活动演示文稿中插入整页 OOXML (.pptx)，元素落地为真实原生对象。"""

    @property
    def name(self) -> str:
        return "ppt_insert_slides_ooxml"

    @property
    def description(self) -> str:
        return (
            "Insert one or more whole PowerPoint slides from a .pptx (OOXML) package into the active "
            "presentation, as REAL native objects (text boxes, shapes, images, tables, native charts). "
            "Use this for HIGH-FIDELITY whole-slide generation — layouts, line spacing, legends, table "
            "borders, image styles that the fine-grained tools cannot achieve (Office.js capability gaps). "
            "Pass `source_path` (a filesystem path to the .pptx you built offline, e.g. with python-pptx); "
            "the Server reads and encodes it — never pass base64 inline. "
            "Inserted elements are addressable by shape id: enumerate them with ppt_get_slide_elements and "
            "edit them afterwards with ppt_update_element / ppt_update_text_box / ppt_update_table_*. "
            "NOTE: when replaceSlideId + finalSlideIndex are given, the composite insert→delete→move runs "
            "sequentially and is NOT atomic (no rollback); on partial failure the error carries "
            "{stage, partiallyApplied, createdSlideId}. Requires the document open in PowerPoint via the "
            "Add-In (PowerPointApi 1.8); otherwise returns a 3003 'close/open the document' style error."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptInsertSlidesOoxmlInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "insert:slidesOoxml"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptInsertSlidesOoxmlInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Read the .pptx file → base64 (Server-side) → emit ppt:insert:slidesOoxml.

        Overrides the base flow only to translate ``source_path`` into the wire ``base64``;
        connection routing, response unwrapping and error formatting are delegated to
        ``workspace.execute()`` (the same path every other write tool uses).
        """
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, PptInsertSlidesOoxmlInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 读取 .pptx 文件 → base64 (整个 base64 留在 Server 端，绝不进模型上下文)
        src = Path(validated.source_path)
        if not src.is_file():
            return {"success": False, "error": f"3004: source_path not found or not a file: {src}"}
        try:
            raw = src.read_bytes()
        except OSError as e:
            return {"success": False, "error": f"3004: failed to read source_path: {e}"}
        if not raw:
            return {"success": False, "error": "3004: source_path is empty"}
        wire_base64 = base64lib.b64encode(raw).decode("ascii")

        # 3. 组装 wire 参数 (camelCase via by_alias)，去掉 source_path，注入 base64
        params = validated.model_dump(by_alias=True, exclude_none=True)
        params.pop("source_path", None)
        params.pop("document_uri", None)
        params["base64"] = wire_base64

        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={"document_uri": validated.document_uri, **params},
        )

        # 4. 执行 (序列化同文档写操作，避免与图表往返/其他整页写交错)
        try:
            async with document_lock_manager.acquire(validated.document_uri):
                obs = await self.workspace.execute(action)
        except Exception as e:  # noqa: BLE001 - surface unexpected I/O as 3004
            logger.exception(f"工具执行失败 | Tool execution failed: {self.name}")
            return {"success": False, "error": f"3004: {e}"}

        # 4.5 活动追踪 + 资源通知
        if obs.success:
            self.workspace.update_last_activity(
                document_uri=validated.document_uri,
                tool_name=self.name,
                result_data=obs.data,
            )
            self.workspace.notify_resource_updated(["window://office4ai/ppt", "window://office4ai"])

        # 5. 格式化返回
        return self.format_result(obs)
