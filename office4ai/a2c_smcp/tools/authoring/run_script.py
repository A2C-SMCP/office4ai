"""office_run_script MCP Tool (milestone #4 · S1 — authoring 脚本运行时).

standalone 工具：不走 Socket.IO / Add-In 连接，在服务端软沙箱子进程里执行 LLM 提交的
Python 脚本（python-docx / pptx / openpyxl / lxml / docxtpl），落盘生成或修改 Office 文件，
回传结构化结果。无 Add-In 连接也可用，故 ``requires_connection=False`` 常驻。
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.office.authoring import run_script
from office4ai.office.authoring.runtime import ScriptResult


class OfficeRunScriptInput(BaseModel):
    """MCP 输入模型：提交一段 Python 脚本到沙箱运行时执行。"""

    script: str = Field(
        ...,
        description=(
            "Inline Python source to execute in the sandbox. The script is fully autonomous: it decides "
            "which file(s) to create/edit, whether to save, and every action. Write products INTO the "
            "provided working directory (its cwd). Allowed imports: python-docx (docx), python-pptx (pptx), "
            "openpyxl, lxml, docxtpl, jinja2, PIL, and a safe stdlib subset. Network access is disabled; "
            "writes outside the working directory are blocked; runaway loops are killed on timeout."
        ),
    )
    work_dir: str | None = Field(
        default=None,
        description=(
            "Working directory the script runs in (== cwd == filesystem write whitelist). Products land here "
            "and this path is returned as `path`. Defaults to a fresh managed directory when omitted."
        ),
    )
    template_uri: str | None = Field(
        default=None,
        description="Optional template file path/URI the script reads (e.g. a .dotx/.potx/.xltx). Validated to exist.",
    )
    timeout: float | None = Field(
        default=None,
        description="Optional wall-clock timeout in seconds (default 60). The process group is killed when exceeded.",
    )


class OfficeRunScriptTool(BaseTool):
    """在软沙箱中运行 LLM 提交的脚本以创建/编辑 Office 文件（无需 Add-In 连接）。

    standalone 工具常驻：``category='authoring'`` → ``BaseTool.requires_connection`` 默认 ``False``，
    W4a 的动态工具收敛据此不过滤本工具（无 Add-In 连接也暴露）。
    """

    @property
    def name(self) -> str:
        return "office_run_script"

    @property
    def description(self) -> str:
        return (
            "Run an inline Python script in a stable server-side sandbox to create or edit Office files "
            "(Word/PowerPoint/Excel) WITHOUT requiring an Add-In connection. The sandbox ships pinned "
            "python-docx / python-pptx / openpyxl / lxml / docxtpl so scripts never manage their own "
            "environment. The script is fully autonomous — it chooses filenames, whether to save, and all "
            "operations; write your outputs into the working directory given by `path`. "
            "Sandbox guardrails (best-effort): filesystem writes are confined to the working directory, "
            "network is disabled, disallowed imports are rejected, and runaway scripts are killed on timeout. "
            "Returns {ok, path, summary, logs, stderr}: ok=whether the script ran successfully, "
            "path=the working directory (list it to find products), summary=manifest of files created/modified, "
            "logs=captured stdout, stderr=captured stderr (sandbox-block reasons and tracebacks appear here). "
            "Do NOT confuse this with the ONLINE escape hatches `word_run_script` / `ppt_run_script` / "
            "`excel_run_script`: those relay raw Office.js JavaScript into a LIVE Add-In to act on the "
            "currently-open document. THIS tool is the OFFLINE channel — it runs Python on disk and needs no "
            "Add-In. If a document is open in a connected Add-In and you want to edit that live document, use "
            "the matching online `*_run_script` tool for that host instead."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return OfficeRunScriptInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel", "authoring"]:
        return "authoring"

    @property
    def event_name(self) -> str:
        # standalone —— 不映射任何 Socket.IO 事件（execute 被完全 override）
        return "run:script"

    @property
    def input_model(self) -> type[BaseModel]:
        return OfficeRunScriptInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Override：绕过 workspace/Socket.IO，直接调 authoring 沙箱运行时。"""
        try:
            validated = self.validate_input(arguments, OfficeRunScriptInput)
        except ValueError as e:
            # 保持 summary 形状与成功路径一致（{produced, count, truncated}），避免消费者 KeyError
            return ScriptResult(ok=False, path="", stderr=f"[runtime] invalid input: {e}").to_contract()

        result = await run_script(
            validated.script,
            work_dir=validated.work_dir,
            template_uri=validated.template_uri,
            timeout=validated.timeout,
        )
        return result.to_contract()
