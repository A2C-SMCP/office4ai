"""excel_run_script MCP Tool —— 在线 Office.js 脚本执行逃生舱（纯中转，issue #87）."""

from __future__ import annotations

from typing import ClassVar, Literal

from office4ai.a2c_smcp.tools.run_script_base import RunScriptToolBase


class ExcelRunScriptTool(RunScriptToolBase):
    """把一段 Office.js JS 中转给在线 Excel Add-In 执行（逃生舱，优先用 typed excel_* 工具）。"""

    _HOST: ClassVar[str] = "Excel"
    _CONTEXT_CLASS: ClassVar[str] = "Excel.RequestContext"
    _EXAMPLE: ClassVar[str] = (
        "`const rng = context.workbook.getSelectedRange(); rng.load('address'); "
        "await context.sync(); return rng.address;`"
    )

    @property
    def name(self) -> str:
        return "excel_run_script"

    @property
    def category(self) -> Literal["word", "ppt", "excel", "authoring"]:
        return "excel"
