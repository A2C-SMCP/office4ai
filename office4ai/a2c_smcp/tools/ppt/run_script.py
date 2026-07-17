"""ppt_run_script MCP Tool —— 在线 Office.js 脚本执行逃生舱（纯中转，issue #87）."""

from __future__ import annotations

from typing import ClassVar, Literal

from office4ai.a2c_smcp.tools.run_script_base import RunScriptToolBase


class PptRunScriptTool(RunScriptToolBase):
    """把一段 Office.js JS 中转给在线 PowerPoint Add-In 执行（逃生舱，优先用 typed ppt_* 工具）。"""

    _HOST: ClassVar[str] = "PowerPoint"
    _CONTEXT_CLASS: ClassVar[str] = "PowerPoint.RequestContext"
    _EXAMPLE: ClassVar[str] = (
        "`const slides = context.presentation.slides; slides.load('items'); "
        "await context.sync(); return slides.items.length;`"
    )

    @property
    def name(self) -> str:
        return "ppt_run_script"

    @property
    def category(self) -> Literal["word", "ppt", "excel", "authoring"]:
        return "ppt"
