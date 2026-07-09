"""共享 OOXML 底层工具：命名空间、qname、包(zip)读写、content-type 常量。

Shared OOXML primitives: namespaces, qnames, package (zip) IO, content-type constants.

authoring helper 的多数原语走「zip + lxml」而非对象模型——因为真实企业模板里的
SDT 内容控件 / 条件格式 / 图表等扩展，在 python-docx / openpyxl 的 load→save round-trip 中
会被丢弃或损坏（实测结论，见规格 §4）。本模块集中封装这条路径上重复出现的样板代码：
把整个包读进内存字典、按 ``ZipInfo`` 原顺序回写、以及命名空间/序列化的统一约定。
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

# ── OOXML 命名空间 | namespaces ─────────────────────────────────────────────
NS: dict[str, str] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xml": "http://www.w3.org/XML/1998/namespace",
}


def qn(prefix_tag: str) -> str:
    """把 ``'w:sdt'`` 形式的前缀名展开成 lxml 的 Clark notation ``'{ns}sdt'``。

    Expand a ``'prefix:tag'`` name to lxml Clark notation. 未知前缀抛 ``KeyError``。
    """
    prefix, sep, tag = prefix_tag.partition(":")
    if not sep:  # 无前缀 -> 视为无命名空间的裸标签 | no prefix -> bare tag
        return prefix
    try:
        return f"{{{NS[prefix]}}}{tag}"
    except KeyError as exc:
        raise KeyError(f"未知 OOXML 命名空间前缀 | unknown namespace prefix: {prefix!r}") from exc


#: ``xml:space`` 属性的 Clark notation，用于保住占位文本的首尾空格。
XML_SPACE: str = qn("xml:space")

# ── 模板 <-> 文档 的 content-type 常量 | template<->document content types ──
#: PowerPoint 模板 / 文档主部件 content-type（``instantiate_from_template`` 的 potx swap 用）。
CT_POTX = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
CT_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
#: Word / Excel 模板主部件 content-type（走 LibreOffice 实例化，此处仅备查/断言用）。
CT_DOTX = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
CT_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
CT_XLTX = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"


def read_package(path: str | Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    """把 OOXML 包内全部部件读进内存，返回 ``(infolist, {部件名: 字节})``。

    保留原始 ``ZipInfo`` 顺序，便于 :func:`write_package` 原样回写（顺序影响某些消费者）。
    先整体读入再关闭句柄，允许 ``path`` 与写出目标相同（原地编辑）。
    """
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        items = {info.filename: zin.read(info.filename) for info in infos}
    return infos, items


def write_package(path: str | Path, infos: list[zipfile.ZipInfo], items: dict[str, bytes]) -> None:
    """按原 ``ZipInfo`` 顺序把内存部件回写成 OOXML 包（``ZIP_DEFLATED``）。

    仅回写 ``infos`` 中列出的部件（S2 原语只改既有部件、不新增部件）。
    """
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, items[info.filename])


def parse_xml(data: bytes) -> etree._Element:
    """把部件字节解析成 lxml 元素 | Parse part bytes into an lxml element."""
    return etree.fromstring(data)


def serialize_xml(element: etree._Element) -> bytes:
    """序列化成带 XML 声明、``standalone`` 的 UTF-8 字节，与 Office 部件的写法一致。"""
    data: bytes = etree.tostring(element, xml_declaration=True, encoding="UTF-8", standalone=True)
    return data
