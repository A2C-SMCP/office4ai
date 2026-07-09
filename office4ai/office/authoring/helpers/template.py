"""模板实例化原语：按平台自动选「LibreOffice 无头」或「content-type 互换」把模板落成文档。

Template instantiation: turn a ``.dotx``/``.potx``/``.xltx`` template into a working
``.docx``/``.pptx``/``.xlsx`` document, auto-selecting the faithful strategy per platform.

合流自三份参考脚本 PART3。模板与文档本质只差包内一个 content-type 标记，Office 点「基于模板新建」
即把它落成文档。按规格 D3，本原语按平台选策略：

- **``.potx`` → content-type 互换**：``.potx`` 本身就是合法演示文稿，只翻一个标记即成 ``.pptx``，
  逐字节保真、占位符 idx/位置/主题样式全不变。（LibreOffice 转 pptx 会把占位符 idx 归 0、重排位置，
  导致填充错位，故 PPT **不**走它。）
- **``.dotx`` / ``.xltx`` → LibreOffice 无头**：Word/Excel 的 LibreOffice 实例化忠实
  （SDT / 条件格式 / 结构化表格实测全部幸存），故沿用。
"""

from __future__ import annotations

import os
from pathlib import Path

from . import _soffice
from ._ooxml import CT_POTX, CT_PPTX, read_package, write_package
from .errors import TemplateFormatError

#: 走 LibreOffice 实例化的模板扩展名 -> 目标文档扩展名。
_SOFFICE_TARGETS = {".dotx": "docx", ".xltx": "xlsx"}


def instantiate_from_template(
    template_path: str | Path,
    output_path: str | Path,
    *,
    soffice: str | None = None,
    timeout: float = 120.0,
) -> Path:
    """把模板实例化为成品文档，写到 ``output_path`` 并返回其路径（等价 Office「基于模板新建」）。

    Instantiate a template into a document at ``output_path``.

    按 ``template_path`` 的扩展名分派：``.potx`` 走 content-type 互换；``.dotx``/``.xltx`` 走
    LibreOffice 无头转换。``output_path`` 应带对应文档扩展名（``.pptx``/``.docx``/``.xlsx``）。
    未知扩展名抛 :class:`TemplateFormatError`。
    """
    template = Path(template_path)
    output = Path(output_path)
    ext = template.suffix.lower()

    if ext == ".potx":
        return _instantiate_potx(template, output)
    if ext in _SOFFICE_TARGETS:
        return _instantiate_via_soffice(template, output, _SOFFICE_TARGETS[ext], soffice, timeout)
    raise TemplateFormatError(
        f"不支持的模板扩展名 | unsupported template extension: {ext!r}（支持 | supported: .dotx, .potx, .xltx）",
    )


def _instantiate_potx(template: Path, output: Path) -> Path:
    """``.potx`` -> ``.pptx``：只翻 content-type 标记，其余字节原样保留（最忠实）。"""
    infos, items = read_package(template)
    content_types = items["[Content_Types].xml"]
    if CT_POTX.encode() not in content_types:
        raise TemplateFormatError(
            f"不是有效的 .potx 模板（缺 presentation template content-type）| not a valid .potx: {template.name}",
        )
    items["[Content_Types].xml"] = content_types.replace(CT_POTX.encode(), CT_PPTX.encode())
    output.parent.mkdir(parents=True, exist_ok=True)
    write_package(output, infos, items)
    return output


def _instantiate_via_soffice(
    template: Path,
    output: Path,
    target_ext: str,
    soffice: str | None,
    timeout: float,
) -> Path:
    """``.dotx``/``.xltx`` -> 文档：LibreOffice 无头转换后重命名到 ``output``。"""
    output.parent.mkdir(parents=True, exist_ok=True)
    produced = _soffice.convert(template, output.parent, target_ext, soffice=soffice, timeout=timeout)
    if produced.resolve() != output.resolve():
        os.replace(produced, output)
    return output
