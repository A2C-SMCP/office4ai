"""共享的 ``{ns}_run_script`` 在线逃生舱基类与输入模型（issue #87 / oasp#18）.

三命名空间（word/ppt/excel）各暴露一个独立工具 ``{ns}_run_script``，把一段 JS 源码**纯中转**
给对应 Office Add-In，注入宿主 ``RequestContext`` 后按 Office.js 语义执行，回传返回值 + 日志。
三者仅 ``name`` / ``category`` / 宿主上下文（描述用）不同，共享同一输入模型、注解、超时派生与
执行流（走 ``BaseTool.execute`` 默认路径 → ``{category}:run:script`` 事件）。

与**离线** ``office_run_script``（服务端 Python/OOXML 沙箱、无需 Add-In、落盘生成文件）
两通道独立、互指防混淆：本组是**在线** Office.js 直连当前打开的文档。
"""

from __future__ import annotations

from typing import Any, ClassVar

from mcp.types import ToolAnnotations
from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool

#: 「脚本执行」档默认超时（毫秒）——与 OASP conventions §run-script 一致（缺省 timeoutMs）
DEFAULT_SCRIPT_TIMEOUT_MS = 60000

#: Server ack 超时相对脚本超时的宽限（毫秒）。server_timeout = (timeoutMs ?? 60000) + GRACE，
#: 使 Add-In 成为脚本时长执法者（超时回 1002），Server ack 超时仅失联兜底、不先于脚本挂断。
SERVER_TIMEOUT_GRACE_MS = 10000


class RunScriptInput(BaseModel):
    """MCP 输入模型：向在线 Office Add-In 下发一段 JS 脚本执行（三命名空间共享）。"""

    document_uri: str = Field(
        ...,
        description="Target document URI of the live Add-In session (e.g. file:///path/to/report.docx)",
    )
    script: str = Field(
        ...,
        description=(
            "Raw Office.js JavaScript to run INSIDE the live Add-In. Executed as an async function body: "
            "use `return` to hand back a JSON-serializable value, read injected params via the global `args`, "
            "and `console.log(...)` for logs. The host `RequestContext` is injected as the global `context` "
            "(call `await context.sync()` to flush). Non-sandboxed: full Office.js is available and the script "
            "may `fetch` arbitrary hosts."
        ),
    )
    args: dict[str, Any] | None = Field(
        default=None,
        description="Optional JSON-serializable params injected into the script as the global `args`.",
    )
    timeout_ms: int | None = Field(
        default=None,
        description=(
            "Optional script-execution timeout in milliseconds (default 60000, no hard cap). Enforced by the "
            "Add-In (returns error 1002 on timeout); the server ack waits slightly longer as a liveness backstop."
        ),
    )


class RunScriptToolBase(BaseTool):
    """在线 ``{ns}:run:script`` 逃生舱工具基类（纯中转）。

    子类只需声明 ``name`` / ``category`` 及三个宿主描述常量
    （``_HOST`` / ``_CONTEXT_CLASS`` / ``_EXAMPLE``）；其余（输入模型、JSON Schema、
    ``event_name='run:script'``、注解、server_timeout 派生）全部在此收敛。
    ``requires_connection`` 经 ``category`` 派生默认 ``True`` → 随 Add-In 连接经 W4a 收敛。
    """

    #: 宿主名（"Word" / "PowerPoint" / "Excel"）——子类设置，仅用于描述文案
    _HOST: ClassVar[str]
    #: 注入脚本的宿主上下文类型名（"Word.RequestContext" 等）——子类设置
    _CONTEXT_CLASS: ClassVar[str]
    #: 一行宿主示例片段——子类设置
    _EXAMPLE: ClassVar[str]

    @property
    def input_schema(self) -> dict[str, Any]:
        return RunScriptInput.model_json_schema()

    @property
    def event_name(self) -> str:
        return "run:script"

    @property
    def input_model(self) -> type[BaseModel]:
        return RunScriptInput

    @property
    def annotations(self) -> ToolAnnotations | None:
        # destructiveHint：脚本可写文档且不保证原子性/回滚；openWorldHint：可 fetch 任意第三方。
        return ToolAnnotations(destructiveHint=True, openWorldHint=True)

    @property
    def description(self) -> str:
        return (
            # 首句锚定逃生舱定位 + 强制「优先 typed」——仅两准入条件（能力缺口 / confirmed bug）
            f"ESCAPE HATCH — ALWAYS prefer the typed {self._HOST} tools. Reach for this ONLY when "
            f"(1) the capability you need has no typed {self._HOST} tool at all, or (2) a specific typed "
            f"{self._HOST} tool has a confirmed bug that blocks you. If a typed tool exists and works, use it — "
            "do NOT use run:script for convenience, batching multiple ops, or to avoid learning a typed tool. "
            f"What it does: executes your raw Office.js JavaScript INSIDE the live {self._HOST} Add-In. The "
            f"Add-In injects a host `{self._CONTEXT_CLASS}` as the global `context`; your script runs as an async "
            "function body (use `return` for a JSON-serializable result; read injected params via the global "
            "`args`; `console.log(...)` for logs). "
            "Returns {result, logs, durationMs, logsTruncated}; result<=512KB, logs<=100KB (UTF-8 bytes). "
            f"Example: {self._EXAMPLE} "
            "Do NOT confuse this with `office_run_script`: that is the OFFLINE channel — a server-side "
            "Python/OOXML sandbox that writes files on disk WITHOUT any Add-In connection. THIS tool is the "
            "ONLINE channel — Office.js against the currently-open document. "
            "Caveats: the script is NON-sandboxed (full Office.js + can `fetch` arbitrary hosts) and NOT atomic — "
            "a failure can leave the document in a partial state with no rollback, so read back inside the script "
            "if you must reconcile. `timeout_ms` bounds script execution (default 60s), enforced by the Add-In."
        )

    def derive_server_timeout_ms(self, params: dict[str, Any]) -> int | None:
        """派生 Server ack 超时 = (timeout_ms ?? 60000) + GRACE。

        非法/非正的 ``timeout_ms`` 回退到默认档（校验交由下游/Add-In，此处只保证 ack 兜底合理）。
        """
        requested = params.get("timeout_ms")
        base = requested if isinstance(requested, int) and requested > 0 else DEFAULT_SCRIPT_TIMEOUT_MS
        return base + SERVER_TIMEOUT_GRACE_MS
