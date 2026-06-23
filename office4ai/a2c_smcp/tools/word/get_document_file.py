"""word_get_document_file MCP Tool (OASP /word Draft, v0.3.0, #46-D — whole-document .docx).

Export the whole document as a base64 ``.docx`` package. Companion to
``word_insert_document_file`` for whole-document round-trip — the path that
``word_get_ooxml`` / ``word_insert_ooxml`` cannot serve, because Office.js
``insertOoxml`` rejects a whole-document Flat OPC package (GeneralException).

Context-cost design (office4ai #46, aligned with PPT #40/#42)
============================================================
The exported base64 is decoded to ``dest_path`` on disk; the tool returns only
``{filePath, bytes}`` — the (large) base64 never enters the model's conversation
context. Mirrors the PPT slide-OOXML binary file-handle pattern (#40/#42).
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any, Literal

from loguru import logger
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeAction


class WordGetDocumentFileInput(BaseModel):
    """MCP 输入模型: 导出整篇 .docx 到文件 (不回 inline base64)。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    dest_path: str = Field(
        ...,
        description=(
            "Filesystem path to write the exported whole-document .docx package. "
            "The base64 is decoded to disk here and is NOT returned inline."
        ),
    )

    model_config = {"populate_by_name": True}


class WordGetDocumentFileTool(BaseTool):
    """导出整篇文档当前状态为 .docx 文件，返回句柄 (不回 inline base64)。"""

    @property
    def name(self) -> str:
        return "word_get_document_file"

    @property
    def description(self) -> str:
        return (
            "Export the WHOLE Word document (current live state, including unsaved edits) as a .docx "
            "package written to `dest_path` on disk. Returns {filePath, bytes} — the base64 is NOT "
            "returned inline (it stays on disk to keep it out of the model context). The written .docx "
            "can be passed as `source_path` to word_insert_document_file for whole-document round-trip "
            "editing. This is the whole-document path; for a selection / OOXML fragment use word_get_ooxml. "
            "Requires the document open in Word via the Add-In."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordGetDocumentFileInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "get:documentFile"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordGetDocumentFileInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Emit word:get:documentFile → decode the returned base64 to ``dest_path`` → return the handle."""
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, WordGetDocumentFileInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 触发导出 (走通用 execute 路径: 连接检查 + emit + 解包)
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={"document_uri": validated.document_uri},
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
            dest.write_bytes(base64.b64decode(wire_base64))
        except (OSError, ValueError) as e:
            return {"success": False, "error": f"3004: failed to write dest_path: {e}"}

        # 4. 活动追踪 + 返回句柄 (不含 base64)
        self.workspace.update_last_activity(
            document_uri=validated.document_uri,
            tool_name=self.name,
            result_data={"filePath": str(dest)},
        )
        return {
            "success": True,
            "data": {
                "filePath": str(dest),
                "bytes": dest.stat().st_size,
            },
        }
