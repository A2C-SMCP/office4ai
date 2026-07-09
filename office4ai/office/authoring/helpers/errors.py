"""authoring helper 的异常类型 | Exception types for the authoring helpers.

原语在遇到「模板/数据不匹配、模板格式不认、缺少外部工具」等可预期的错误时抛出这些异常，
方便 SKILL 脚本 / 上层运行时按类型捕获处理。所有异常都继承 :class:`AuthoringHelperError`。
"""

from __future__ import annotations


class AuthoringHelperError(Exception):
    """authoring helper 抛出的所有异常的基类 | Base class for all authoring helper errors."""


class SofficeNotFoundError(AuthoringHelperError):
    """未找到 LibreOffice(soffice) 可执行文件 | LibreOffice (soffice) executable not found.

    ``.dotx`` / ``.xltx`` 的忠实实例化依赖 LibreOffice 无头转换；找不到时抛出。
    """


class SofficeConversionError(AuthoringHelperError):
    """LibreOffice 无头转换失败(非零退出 / 超时 / 未产出目标文件)。"""


class TemplateFormatError(AuthoringHelperError):
    """模板文件的格式/扩展名不被支持 | Unsupported template format or extension."""


class SdtCountMismatchError(AuthoringHelperError):
    """SDT 内容控件数量与提供的值数量不一致(严格模式) | SDT control count != values count."""


class AnchorNotFoundError(AuthoringHelperError):
    """在文档中定位不到请求的锚点(样式 / 命名区域 / 版式) | Requested anchor not found."""
