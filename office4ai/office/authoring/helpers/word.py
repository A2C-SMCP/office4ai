"""Word 原语：SDT 内容控件填充 + 按样式定位锚点。

Word primitives: fill SDT content controls, locate anchors by paragraph style.

提炼自参考脚本 ``gen_word.py``。核心是 :func:`fill_sdt_controls`——Office 精美模板的标题/
小标题常用 **SDT 内容控件**（灰显占位），python-docx/docxtpl 改不动，必须走 zip + lxml 兜底。
本原语同时支持**有序列表**（忠实于参考脚本、可处理重复 alias）与 **alias 字典**（抗模板漂移）两种入参。
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from ._ooxml import XML_SPACE, parse_xml, qn, read_package, serialize_xml, write_package
from .errors import AnchorNotFoundError, SdtCountMismatchError

# SDT 上的占位标记：填入正式内容后须移除，否则文字仍灰显、且会被 Office 当占位符再次清空。
_PLACEHOLDER_MARKS = ("w:showingPlcHdr", "w:temporary")


@dataclass(frozen=True)
class StyleMatch:
    """一处按段落样式命中的锚点 | An anchor matched by paragraph style.

    ``index`` 为该段落在正文全部段落中的文档顺序序号；``text`` 为其纯文本。
    """

    index: int
    text: str


def _iter_text_sdts(doc: etree._Element) -> Iterator[tuple[etree._Element, list[etree._Element]]]:
    """按文档顺序产出所有「含文字」的 SDT 控件及其 ``w:t`` 列表（行内级/段落级统一以 w:sdt 为单位）。"""
    for sdt in doc.iter(qn("w:sdt")):
        content = sdt.find(qn("w:sdtContent"))
        if content is None:
            continue
        texts = list(content.iter(qn("w:t")))
        if texts and "".join(t.text or "" for t in texts).strip():
            yield sdt, texts


def _sdt_alias(sdt: etree._Element) -> str | None:
    """取 SDT 的 alias（无则回退 tag），用于字典模式按名匹配。"""
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is None:
        return None
    for tag in ("w:alias", "w:tag"):
        el = sdt_pr.find(qn(tag))
        if el is not None:
            val: str | None = el.get(qn("w:val"))
            if val:
                return val
    return None


def _fill_one(texts: list[etree._Element], sdt: etree._Element, value: object) -> None:
    """把 value 写进 SDT 的首个 ``w:t``、清空其余，并去掉占位标记（转为正式内容）。"""
    texts[0].text = str(value)
    texts[0].set(XML_SPACE, "preserve")
    for extra in texts[1:]:
        extra.text = ""
    sdt_pr = sdt.find(qn("w:sdtPr"))
    if sdt_pr is not None:
        for mark in _PLACEHOLDER_MARKS:
            for el in sdt_pr.findall(qn(mark)):
                sdt_pr.remove(el)


def fill_sdt_controls(
    docx_path: str | Path,
    values: Sequence[object] | Mapping[str, object],
    *,
    output: str | Path | None = None,
    strict: bool = True,
) -> int:
    """填充 ``.docx`` 内所有「含文字」的 SDT 内容控件，返回填入的控件数。

    Fill every text-bearing SDT content control in a ``.docx``.

    两种 ``values`` 入参：

    - **有序序列**（``list``/``tuple``）：按文档顺序逐个填充；``strict=True`` 时数量须与控件数一致，
      便于及早发现模板/数据错位。适合含**重复 alias** 的真实模板（如首/次版式同名标题）。
    - **映射**（``{alias: value}``）：按 SDT 的 alias/tag 匹配填充；``strict=True`` 时映射中每个键
      都必须命中。同名 alias 只填第一个——需要填重复项请改用有序序列。

    填完会去掉 ``w:showingPlcHdr`` / ``w:temporary`` 占位标记，使文字成为正式内容（不再灰显）。
    ``output=None`` 表示原地写回 ``docx_path``。全程 zip + lxml，不经 python-docx，避免 SDT 结构损坏。

    **范围限制**：当前仅处理正文 ``word/document.xml`` 的 SDT，不含页眉/页脚
    (``header*.xml`` / ``footer*.xml``) 里的控件（后者不填也不计数）。若模板把 SDT 放在页眉/页脚，
    请用 :func:`locate_by_style` 之外的手段另行处理——扩展到页眉/页脚为后续跟进项。
    """
    if isinstance(values, str):  # str 也是 Sequence，会被逐字符拆开——显式拒绝
        raise TypeError("values 不能是 str；请传序列或映射 | values must be a sequence or mapping, not str")

    infos, items = read_package(docx_path)
    doc = parse_xml(items["word/document.xml"])
    sdts = list(_iter_text_sdts(doc))

    if isinstance(values, Mapping):
        filled = _fill_by_alias(sdts, values, strict)
    else:
        seq = list(values)
        if strict and len(sdts) != len(seq):
            raise SdtCountMismatchError(
                f"SDT 控件数({len(sdts)})与提供值数({len(seq)})不一致 | SDT count != values count",
            )
        filled = 0
        # 严格模式已在上面校验等长；非严格模式按较短方截断 | strict already length-checked above
        for (sdt, texts), value in zip(sdts, seq, strict=False):
            _fill_one(texts, sdt, value)
            filled += 1

    items["word/document.xml"] = serialize_xml(doc)
    write_package(output or docx_path, infos, items)
    return filled


def _fill_by_alias(
    sdts: list[tuple[etree._Element, list[etree._Element]]],
    mapping: Mapping[str, object],
    strict: bool,
) -> int:
    """字典模式：按 alias/tag 匹配填充；同名只填第一个。"""
    remaining = dict(mapping)
    filled = 0
    for sdt, texts in sdts:
        alias = _sdt_alias(sdt)
        if alias is not None and alias in remaining:
            _fill_one(texts, sdt, remaining.pop(alias))
            filled += 1
    if strict and remaining:
        raise AnchorNotFoundError(
            f"这些 SDT alias 未在文档中找到 | SDT aliases not found: {sorted(remaining)}",
        )
    return filled


def locate_by_style(docx_path: str | Path, style_id: str) -> list[StyleMatch]:
    """定位 ``.docx`` 正文中所有使用指定段落样式(``w:pStyle``)的段落，返回 :class:`StyleMatch` 列表。

    Locate paragraphs whose ``w:pStyle`` equals ``style_id``（如 ``"Heading1"``）。
    W3 抽模板时据此找出结构锚点（把标题段落升级为 SDT / 占位符）。样式 id 区分大小写。
    """
    _, items = read_package(docx_path)
    doc = parse_xml(items["word/document.xml"])
    body = doc.find(qn("w:body"))
    matches: list[StyleMatch] = []
    if body is None:
        return matches
    for idx, para in enumerate(body.iter(qn("w:p"))):
        p_pr = para.find(qn("w:pPr"))
        if p_pr is None:
            continue
        p_style = p_pr.find(qn("w:pStyle"))
        if p_style is not None and p_style.get(qn("w:val")) == style_id:
            text = "".join(t.text or "" for t in para.iter(qn("w:t")))
            matches.append(StyleMatch(index=idx, text=text))
    return matches
