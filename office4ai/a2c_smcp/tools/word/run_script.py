"""word_run_script MCP Tool —— 在线 Office.js 脚本执行逃生舱（纯中转，issue #87）."""

from __future__ import annotations

from typing import ClassVar, Literal

from office4ai.a2c_smcp.tools.run_script_base import RunScriptToolBase


class WordRunScriptTool(RunScriptToolBase):
    """把一段 Office.js JS 中转给在线 Word Add-In 执行（逃生舱，优先用 typed word_* 工具）。"""

    _HOST: ClassVar[str] = "Word"
    _CONTEXT_CLASS: ClassVar[str] = "Word.RequestContext"
    _EXAMPLE: ClassVar[str] = (
        "`context.document.body.insertParagraph('Hello', 'End'); await context.sync(); return 'done';`"
    )

    @property
    def name(self) -> str:
        return "word_run_script"

    @property
    def category(self) -> Literal["word", "ppt", "excel", "authoring"]:
        return "word"
