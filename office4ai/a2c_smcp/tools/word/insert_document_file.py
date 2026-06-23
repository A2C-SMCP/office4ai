"""word_insert_document_file MCP Tool (OASP /word Draft, v0.3.0, #46-D — whole-document .docx).

Insert a whole ``.docx`` package into a Word document. This is the whole-document
round-trip path that ``word_insert_ooxml`` cannot serve: Office.js ``insertOoxml``
rejects a whole-document Flat OPC package (GeneralException), so whole documents go
through base64 ``.docx`` + ``Body/Range.insertFileFromBase64`` instead (mirrors PPT).

Fidelity note (cross-ask F2)
============================
``insertFileFromBase64`` replaces the body *content*; document-level parts (header/
footer, document-level sectPr, docProps, customXml, styles) are NOT guaranteed to
round-trip. Do not assume byte-level symmetry.

Context-cost design (office4ai #46, aligned with PPT #40/#42)
============================================================
The model passes a filesystem path (``source_path``), NOT inline base64. The Server
reads the .docx bytes and base64-encodes them itself before emitting — the base64 on
the wire is a Server→Add-In concern and never enters the model's conversation context.
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any, Literal

from loguru import logger
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.dtos.word import WordInsertLocation, WordOoxmlScope


class WordInsertDocumentFileInput(BaseModel):
    """MCP 输入模型: 整篇 .docx 插入 (文件句柄, 服务端 base64)。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    source_path: str = Field(
        ...,
        description=(
            "Filesystem path to a whole .docx package (typically produced by word_get_document_file, "
            "or authored offline e.g. with python-docx). The Server reads its bytes and base64-encodes "
            "them — DO NOT pass base64 inline. Use this for whole-document insertion/replacement; for a "
            "small OOXML fragment prefer word_insert_ooxml."
        ),
    )
    insert_location: WordInsertLocation = Field(
        ...,
        alias="insertLocation",
        description="Where to insert relative to the anchor range: Replace, Start, or End",
    )
    scope: WordOoxmlScope = Field(
        default="body",
        description='Insert anchor: "body" (whole document body) or "selection" (current selection)',
    )

    model_config = {"populate_by_name": True}


class WordInsertDocumentFileTool(BaseTool):
    """在 Word 文档插入整篇 .docx (base64)，内容经 insertFileFromBase64 落地。"""

    @property
    def name(self) -> str:
        return "word_insert_document_file"

    @property
    def description(self) -> str:
        return (
            "Insert a WHOLE .docx document into a Word document as real native content, via the Add-In's "
            "Body/Range.insertFileFromBase64. This is the whole-document path — use it to round-trip a "
            "whole document exported by word_get_document_file, or to apply a .docx built offline (e.g. "
            "python-docx). Pass `source_path` (a filesystem path to the .docx); the Server reads and "
            "encodes it — never pass base64 inline. `insert_location` is Replace/Start/End relative to the "
            "anchor (`scope` = body or selection; scope=body + Replace replaces the whole body). NOTE: "
            "insertFileFromBase64 replaces body CONTENT — document-level parts (header/footer, document "
            "sectPr, docProps, customXml) are not guaranteed to round-trip, so it is not a pixel-perfect "
            "document rebuild. For a small OOXML fragment use word_insert_ooxml. Requires the document open "
            "in Word via the Add-In."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertDocumentFileInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:documentFile"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertDocumentFileInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Read the .docx file → base64 (Server-side) → emit word:insert:documentFile.

        Overrides the base flow only to translate ``source_path`` into the wire ``base64``;
        connection routing, response unwrapping and error formatting are delegated to
        ``workspace.execute()`` (the same path every other write tool uses).
        """
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, WordInsertDocumentFileInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 读取 .docx 文件 → base64 (整个 base64 留在 Server 端，绝不进模型上下文)
        src = Path(validated.source_path)
        if not src.is_file():
            return {"success": False, "error": f"3004: source_path not found or not a file: {src}"}
        try:
            raw = src.read_bytes()
        except OSError as e:
            return {"success": False, "error": f"3004: failed to read source_path: {e}"}
        if not raw:
            return {"success": False, "error": "3004: source_path is empty"}
        wire_base64 = base64.b64encode(raw).decode("ascii")

        # 3. emit word:insert:documentFile (base64 上 wire；Server↔file↔Add-In)
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={
                "document_uri": validated.document_uri,
                "base64": wire_base64,
                "insertLocation": validated.insert_location,
                "scope": validated.scope,
            },
        )
        try:
            obs = await self.workspace.execute(action)
        except Exception as e:  # noqa: BLE001 - surface unexpected I/O as 3004
            logger.exception(f"工具执行失败 | Tool execution failed: {self.name}")
            return {"success": False, "error": f"3004: {e}"}

        # 4. 活动追踪
        if obs.success:
            self.workspace.update_last_activity(
                document_uri=validated.document_uri,
                tool_name=self.name,
                result_data=obs.data,
            )

        # 5. 格式化返回
        return self.format_result(obs)
