"""word_get_ooxml MCP Tool (OASP /word Draft, v0.3.0 — OOXML fragment round-trip).

Export the live OOXML of the current selection or whole body as a Flat OPC /
WordprocessingML *string*. Companion to ``word_insert_ooxml``: the written file
feeds the latter's ``source_path`` for round-trip editing.

Context-cost design (office4ai #46, aligned with PPT #40/#42)
============================================================
The exported OOXML string is written to ``dest_path`` on disk; the tool returns only
``{scope, filePath, bytes}`` — the (potentially large) XML never enters the model's
conversation context. Unlike PPT (base64 .pptx package), Word OOXML is plain XML
text, so there is no base64 encode/decode — the string is written verbatim.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from loguru import logger
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.dtos.word import WordOoxmlScope


class WordGetOoxmlInput(BaseModel):
    """MCP 输入模型: 导出 OOXML 字符串到文件 (不回 inline)。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    scope: WordOoxmlScope = Field(
        default="selection",
        description=(
            'Export scope: "selection" exports the current selection (falls back to the whole body '
            'when nothing is selected); "body" exports the whole document. The response reports the '
            "effective scope actually used."
        ),
    )
    dest_path: str = Field(
        ...,
        description=(
            "Filesystem path to write the exported OOXML (Flat OPC / WordprocessingML) string. "
            "The XML text is written to disk here and is NOT returned inline."
        ),
    )

    model_config = {"populate_by_name": True}


class WordGetOoxmlTool(BaseTool):
    """导出当前选区/整篇 Body 的 OOXML 字符串为文件，返回句柄 (不回 inline)。"""

    @property
    def name(self) -> str:
        return "word_get_ooxml"

    @property
    def description(self) -> str:
        return (
            "Export the live OOXML (Flat OPC / WordprocessingML string) of the current selection or the "
            "whole document body, written to `dest_path` on disk. Returns {scope, filePath, bytes} — the "
            "OOXML string is NOT returned inline (it stays on disk to keep it out of the model context). "
            "`scope=selection` falls back to the whole body when nothing is selected; the returned `scope` "
            "reports the effective range actually used. The written file is a Flat OPC XML string suitable "
            "as `source_path` for word_insert_ooxml (round-trip editing). Requires the document open in "
            "Word via the Add-In (WordApi 1.1)."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetOoxmlInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:ooxml"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetOoxmlInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Emit word:get:ooxml → write the returned OOXML string to ``dest_path`` → return the handle."""
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, WordGetOoxmlInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 触发导出 (走通用 execute 路径: 连接检查 + emit + 解包)
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={"document_uri": validated.document_uri, "scope": validated.scope},
        )
        try:
            obs = await self.workspace.execute(action)
        except Exception as e:  # noqa: BLE001
            logger.exception(f"工具执行失败 | Tool execution failed: {self.name}")
            return {"success": False, "error": f"3004: {e}"}

        if not obs.success:
            return self.format_result(obs)

        # 3. 把导出的 OOXML 字符串落盘 (不回传 inline)
        ooxml = obs.data.get("ooxml") if isinstance(obs.data, dict) else None
        if not ooxml:
            return {"success": False, "error": "3004: export succeeded but no ooxml in response"}
        dest = Path(validated.dest_path)
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(ooxml, encoding="utf-8")
        except OSError as e:
            return {"success": False, "error": f"3004: failed to write dest_path: {e}"}

        # 4. 活动追踪 + 返回句柄 (不含 ooxml)；scope 回生效值 (空选区回退 body 时为 "body")
        effective_scope = obs.data.get("scope", validated.scope)
        self.workspace.update_last_activity(
            document_uri=validated.document_uri,
            tool_name=self.name,
            result_data={"scope": effective_scope},
        )
        return {
            "success": True,
            "data": {
                "scope": effective_scope,
                "filePath": str(dest),
                "bytes": len(ooxml.encode("utf-8")),
            },
        }
