"""word_insert_ooxml MCP Tool (OASP /word Draft, v0.3.0 — OOXML fragment round-trip).

Insert a Flat OPC / WordprocessingML OOXML fragment into a Word document as REAL
native content (runs, styles, tables, content controls) — the high-fidelity path for
content the fine-grained tools cannot express.

Context-cost design (office4ai #46, aligned with PPT #40/#42)
============================================================
The model passes a **filesystem path** (``source_path``), NOT the inline OOXML. The
Server reads the file's text and sends it over the wire to the Add-In. Unlike PPT
(binary .pptx → base64), Word OOXML is plain Flat OPC XML text, so the file IS the
wire payload verbatim — no base64, no binary .docx, no conversion. Keeping the (large)
XML on disk (Server ↔ file ↔ Add-In) means it never enters the model's context.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from loguru import logger
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.dtos.word import WordInsertLocation, WordOoxmlScope


class WordInsertOoxmlInput(BaseModel):
    """MCP 输入模型: OOXML 片段插入 (文件句柄, Flat OPC 字符串, 非 base64)。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    source_path: str = Field(
        ...,
        description=(
            "Filesystem path to a Flat OPC / WordprocessingML OOXML *string* file (UTF-8 XML text — NOT a "
            "binary .docx and NOT base64). Typically the dest_path written by word_get_ooxml, or authored "
            "offline as Flat OPC. The Server reads its text and sends it on the wire. For small per-element "
            "tweaks prefer word_insert_text / word_replace_text."
        ),
    )
    insert_location: WordInsertLocation = Field(
        ...,
        alias="insertLocation",
        description="Where to insert relative to the anchor range: Replace, Start, or End",
    )
    scope: WordOoxmlScope = Field(
        default="selection",
        description='Insert anchor: "selection" (current selection) or "body" (whole document body)',
    )

    model_config = {"populate_by_name": True}


class WordInsertOoxmlTool(BaseTool):
    """在 Word 文档插入 OOXML 片段 (Flat OPC 字符串)，内容落地为真实原生对象。"""

    @property
    def name(self) -> str:
        return "word_insert_ooxml"

    @property
    def description(self) -> str:
        return (
            "Insert an OOXML *fragment* (Flat OPC / WordprocessingML) into a Word document at the current "
            "selection or whole body, as REAL native content (runs, styles, tables, content controls). "
            "IMPORTANT: `ooxml` must be a FRAGMENT (selection-level content), NOT a whole-document package "
            "— Office.js insertOoxml rejects whole-document OOXML (e.g. the `scope=body` output of "
            "word_get_ooxml) with GeneralException. To insert/round-trip a WHOLE document, use "
            "word_insert_document_file instead. Pass `source_path` — a filesystem path to a Flat OPC XML "
            "*string* file (typically a fragment from word_get_ooxml `scope=selection`, or authored "
            "offline); the Server reads its text and sends it. Do NOT pass a binary .docx or base64 — only "
            "a Flat OPC XML string is accepted. `insert_location` is Replace/Start/End relative to the "
            "anchor (`scope` = selection or body). Malformed OOXML is rejected cleanly by Word without "
            "corrupting the document. For small per-element edits prefer word_insert_text / "
            "word_replace_text. Requires the document open in Word via the Add-In (WordApi 1.1)."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertOoxmlInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:ooxml"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertOoxmlInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Read the Flat OPC XML file → emit word:insert:ooxml with the string on the wire.

        Overrides the base flow only to translate ``source_path`` into the wire ``ooxml``;
        connection routing, response unwrapping and error formatting are delegated to
        ``workspace.execute()`` (the same path every other write tool uses).
        """
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, WordInsertOoxmlInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 读取 Flat OPC XML 字符串 (整个字符串留在 Server 端，绝不进模型上下文)
        src = Path(validated.source_path)
        if not src.is_file():
            return {"success": False, "error": f"3004: source_path not found or not a file: {src}"}
        try:
            ooxml = src.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as e:
            return {"success": False, "error": f"3004: failed to read source_path: {e}"}
        if not ooxml.strip():
            return {"success": False, "error": "3004: source_path is empty"}

        # 3. emit word:insert:ooxml (ooxml 字符串上 wire；Server↔file↔Add-In)
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={
                "document_uri": validated.document_uri,
                "ooxml": ooxml,
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
