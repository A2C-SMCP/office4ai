# filename: __init__.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""MCP 资源集合 | MCP resources."""

from office4ai.a2c_smcp.resources.base import BaseResource
from office4ai.a2c_smcp.resources.ppt_window import PptWindowResource
from office4ai.a2c_smcp.resources.skill import (
    DEFAULT_SKILL_HOST,
    SkillResource,
    SkillResourceError,
    discover_skill_resources,
)
from office4ai.a2c_smcp.resources.window import WindowResource
from office4ai.a2c_smcp.resources.word_window import WordWindowResource

__all__ = [
    "DEFAULT_SKILL_HOST",
    "BaseResource",
    "PptWindowResource",
    "SkillResource",
    "SkillResourceError",
    "WindowResource",
    "WordWindowResource",
    "discover_skill_resources",
]
