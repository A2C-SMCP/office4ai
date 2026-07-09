"""脚本化文档流水线的 authoring 子系统 | Authoring subsystem for the scripted document pipeline.

milestone #4 的地基，含两块：

- ``runtime`` (S1 / #57)：跨平台软沙箱运行时 + soffice 探测；上层 ``office_run_script`` MCP 工具建于此。
- ``helpers`` (S2 / #58)：一组**可 import 的 OOXML 原语**（SDT 填充 / 保图表改单元格 / 新建 slideLayout /
  真实模板实例化），提炼自参考脚本 ``gen_word/ppt/excel.py``，**随 SKILL 的 ``scripts/`` 分发、
  渐进式披露、不做 MCP 工具**（规格 D3 决策记录，避免工具膨胀）。

The ``helpers`` package is a plain importable library on purpose: SKILL scripts import those
primitives rather than the MCP server exposing them as tools.
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
