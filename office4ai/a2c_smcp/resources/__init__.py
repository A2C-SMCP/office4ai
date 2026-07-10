# filename: __init__.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""MCP 资源集合 | MCP resources."""

from office4ai.a2c_smcp.resources.base import BaseResource
from office4ai.a2c_smcp.resources.per_file_window import (
    ExcelFileWindowResource,
    PerFileWindowResource,
    PptFileWindowResource,
    WordFileWindowResource,
    create_per_file_window,
    per_file_window_base_uri,
)
from office4ai.a2c_smcp.resources.skill import (
    DEFAULT_SKILL_HOST,
    SkillResource,
    SkillResourceError,
    discover_skill_resources,
)
from office4ai.a2c_smcp.resources.window import WindowResource

__all__ = [
    "DEFAULT_SKILL_HOST",
    "BaseResource",
    "ExcelFileWindowResource",
    "PerFileWindowResource",
    "PptFileWindowResource",
    "SkillResource",
    "SkillResourceError",
    "WindowResource",
    "WordFileWindowResource",
    "create_per_file_window",
    "discover_skill_resources",
    "per_file_window_base_uri",
]
