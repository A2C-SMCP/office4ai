"""authoring 运行时子包 | authoring runtime subpackage.

milestone #4 · S1 —— 脚本化文档流水线的地基：跨平台软沙箱运行时 + soffice 探测。
上层 ``office_run_script`` MCP 工具与 S2 helper 原语库都建于此。
"""

from office4ai.office.authoring.runtime import (
    DEFAULT_ALLOWED_IMPORTS,
    DEFAULT_TIMEOUT_SECONDS,
    ScriptResult,
    detect_soffice,
    log_soffice_status,
    run_script,
)

__all__ = [
    "DEFAULT_ALLOWED_IMPORTS",
    "DEFAULT_TIMEOUT_SECONDS",
    "ScriptResult",
    "detect_soffice",
    "log_soffice_status",
    "run_script",
]
