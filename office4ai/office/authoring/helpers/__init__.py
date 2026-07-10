"""authoring helper 原语库（S2 / issue #58）——可 import、随 SKILL 分发、渐进式披露。

Authoring helper primitives: an importable OOXML library distributed with SKILL ``scripts/``.

提炼自参考脚本 ``gen_word/ppt/excel.py``，把企业模板复用的高价值能力沉淀为稳定原语：

- **Word**：:func:`fill_sdt_controls`（填 SDT 内容控件）、:func:`locate_by_style`（按段落样式定位锚点）、
  :func:`wrap_in_sdt`（W3 抽取：把段落升级为命名 SDT 占位）、:func:`replace_runs_with_token`
  （W3 抽取：实值→``{{token}}``，跨 run 合并）
- **Excel**：:func:`fill_cells_lxml`（保图表改单元格 + 删公式缓存逼重算）、
  :func:`locate_by_named_range`（定位命名区域）
- **PPT**：:func:`create_custom_layout`（往母版新建自定义版式）、:func:`locate_by_master`（按名定位版式）、
  :func:`promote_slide_to_layout`（W3 抽取：既有幻灯片→可复用版式）
- **跨平台**：:func:`instantiate_from_template`（``.potx`` content-type 互换 / ``.dotx``·``.xltx``
  LibreOffice 无头实例化）

用法示例::

    from office4ai.office.authoring.helpers import fill_sdt_controls, instantiate_from_template

    instantiate_from_template("brand.potx", "deck.pptx")
    fill_sdt_controls("report.docx", {"标题": "2026 年度报告"})
"""

from __future__ import annotations

from ._soffice import find_soffice
from .errors import (
    AnchorNotFoundError,
    AuthoringHelperError,
    ExternalMediaError,
    SdtCountMismatchError,
    SofficeConversionError,
    SofficeNotFoundError,
    TemplateFormatError,
)
from .excel import NamedRange, fill_cells_lxml, locate_by_named_range
from .ppt import create_custom_layout, locate_by_master, promote_slide_to_layout
from .template import instantiate_from_template
from .word import StyleMatch, fill_sdt_controls, locate_by_style, replace_runs_with_token, wrap_in_sdt

__all__ = [
    # Word
    "fill_sdt_controls",
    "locate_by_style",
    "wrap_in_sdt",
    "replace_runs_with_token",
    "StyleMatch",
    # Excel
    "fill_cells_lxml",
    "locate_by_named_range",
    "NamedRange",
    # PPT
    "create_custom_layout",
    "locate_by_master",
    "promote_slide_to_layout",
    # 跨平台 | cross-platform
    "instantiate_from_template",
    "find_soffice",
    # 异常 | errors
    "AuthoringHelperError",
    "SofficeNotFoundError",
    "SofficeConversionError",
    "TemplateFormatError",
    "SdtCountMismatchError",
    "AnchorNotFoundError",
    "ExternalMediaError",
]
