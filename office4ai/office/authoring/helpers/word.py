"""Word 原语：SDT 内容控件填充/抽取 + 按样式定位锚点 + 跨 run 令牌替换。

Word primitives: fill/extract SDT content controls, locate anchors by paragraph style,
run-splitting-aware token replacement.

提炼自参考脚本 ``gen_word.py``。核心是 :func:`fill_sdt_controls`——Office 精美模板的标题/
小标题常用 **SDT 内容控件**（灰显占位），python-docx/docxtpl 改不动，必须走 zip + lxml 兜底。
本原语同时支持**有序列表**（忠实于参考脚本、可处理重复 alias）与 **alias 字典**（抗模板漂移）两种入参。

**W3 抽取方向**（参考文件 → 模板文件）：:func:`wrap_in_sdt` 把定位到的段落升级为命名 SDT 占位框
（与 :func:`fill_sdt_controls` 闭环）；:func:`replace_runs_with_token` 把具体实值换成 ``{{token}}``
（供 docxtpl 回填）。两者都**消解 run-splitting**——真实文档里一个词常被拆进多个 run。
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

    返回的 :attr:`StyleMatch.index` 与 :func:`wrap_in_sdt` 的 ``para_index`` 同一坐标系
    （正文全部 ``w:p`` 的文档顺序序号），可直接把定位结果喂给 :func:`wrap_in_sdt`。
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


# ── W3 抽取方向：把「实值」升级为「模板占位」 | extract direction: value -> placeholder ──


def _set_val(parent: etree._Element, tag: str, val: str) -> etree._Element:
    """在 ``parent`` 下建一个带 ``w:val`` 属性的子元素（SDT 属性的通用写法）。"""
    el = etree.SubElement(parent, qn(tag))
    el.set(qn("w:val"), val)
    return el


def wrap_in_sdt(
    docx_path: str | Path,
    para_index: int,
    alias: str,
    *,
    tag: str | None = None,
    placeholder: str | None = None,
    sdt_id: int | None = None,
    output: str | Path | None = None,
) -> None:
    """把正文中第 ``para_index`` 个段落**原地升级**成一个带 ``alias`` 的 SDT 内容控件（占位框）。

    Wrap the ``para_index``-th body paragraph into an SDT content control (a named placeholder).

    这是 W3 抽模板的「结构锚点」方向：先用 :func:`locate_by_style` 找到标题/字段所在段落，再用本函数
    把它升级为**可复用的命名占位框**——升级后的模板可被 :func:`fill_sdt_controls` 按 ``alias`` 精准回填，
    形成「抽取(wrap) ↔ 填充(fill)」闭环。

    - ``para_index`` 与 :attr:`StyleMatch.index` 同坐标系（正文全部 ``w:p`` 的文档顺序序号）。
    - ``tag`` 缺省等于 ``alias``（``fill_sdt_controls`` 的字典模式按 alias/tag 匹配，两者一致最省心）。
    - ``placeholder`` 非 ``None`` 时，把段落的 run 文本**归一为单个 run** 承载该占位文本——顺带**消解
      run-splitting**（真实值常被 Word 拆进多个 run，若不归一，回填/展示都可能错位）；``None`` 则保留段落
      原有文本作为占位内容。控件带 ``w:showingPlcHdr`` 灰显标记，回填时 ``fill_sdt_controls`` 会自动剥离。
    - ``output=None`` 表示原地写回。全程 zip + lxml，不经 python-docx，避免既有 SDT/结构损坏。

    ``para_index`` 越界或文档无 ``w:body`` 时抛 :class:`AnchorNotFoundError`。
    """
    infos, items = read_package(docx_path)
    doc = parse_xml(items["word/document.xml"])
    body = doc.find(qn("w:body"))
    if body is None:
        raise AnchorNotFoundError("文档缺少 w:body，无法定位段落 | document has no w:body")
    paras = list(body.iter(qn("w:p")))
    if not 0 <= para_index < len(paras):
        raise AnchorNotFoundError(
            f"段落索引越界 para_index={para_index}；正文共 {len(paras)} 段 | paragraph index out of range",
        )
    para = paras[para_index]
    parent = para.getparent()
    if parent is None:
        raise AnchorNotFoundError("目标段落无父节点，无法包裹 | target paragraph has no parent")

    if placeholder is not None:
        texts = list(para.iter(qn("w:t")))
        if texts:
            texts[0].text = placeholder
            texts[0].set(XML_SPACE, "preserve")
            for extra in texts[1:]:  # 归一：多 run 合成一个，消解 run-splitting
                extra.text = ""
        else:  # 段落无 run（如空标题占位）——补一个承载占位文本
            run = etree.SubElement(para, qn("w:r"))
            text_el = etree.SubElement(run, qn("w:t"))
            text_el.text = placeholder
            text_el.set(XML_SPACE, "preserve")

    sdt = etree.Element(qn("w:sdt"))
    sdt_pr = etree.SubElement(sdt, qn("w:sdtPr"))
    _set_val(sdt_pr, "w:alias", alias)
    _set_val(sdt_pr, "w:tag", tag if tag is not None else alias)
    if sdt_id is not None:
        _set_val(sdt_pr, "w:id", str(sdt_id))
    etree.SubElement(sdt_pr, qn("w:showingPlcHdr"))
    content = etree.SubElement(sdt, qn("w:sdtContent"))

    idx = parent.index(para)
    parent.remove(para)
    content.append(para)
    parent.insert(idx, sdt)

    items["word/document.xml"] = serialize_xml(doc)
    write_package(output or docx_path, infos, items)


def _run_ranges(t_nodes: list[etree._Element]) -> tuple[list[tuple[etree._Element, int, int, str]], str]:
    """给一组 ``w:t`` 建立 ``[(节点, 起, 止, 文本)]`` 的字符区间表 + 拼接全文。"""
    ranges: list[tuple[etree._Element, int, int, str]] = []
    cursor = 0
    for node in t_nodes:
        text = node.text or ""
        ranges.append((node, cursor, cursor + len(text), text))
        cursor += len(text)
    return ranges, "".join(r[3] for r in ranges)


def _replace_across_runs(t_nodes: list[etree._Element], old: str, new: str) -> int:
    """在一段的 ``w:t`` 序列里把 ``old`` 全部替换为 ``new``，跨 run 合并，返回替换次数。

    真实文档里一个词常被拆进多个 run（``20|26``）；本函数按**段内拼接全文**定位，再把 ``new`` 整段
    落进匹配起点所在的那个 run、清空被覆盖的其余部分——保证 ``new``（如 ``{{title}}``）落在**单个 run**、
    不被 Word 再次拆断。``search_from`` 越过刚插入的 ``new``，避免 ``new`` 含 ``old`` 时的死循环/重复替换。
    """
    count = 0
    search_from = 0
    while True:
        ranges, full = _run_ranges(t_nodes)
        pos = full.find(old, search_from)
        if pos < 0:
            return count
        end = pos + len(old)
        for node, a, b, text in ranges:
            if b <= pos or a >= end:  # 与匹配区间无交集，原样保留
                continue
            prefix = text[: max(0, pos - a)]  # 本 run 中匹配起点之前的字符
            suffix = text[end - a :] if end < b else ""  # 本 run 中匹配终点之后的字符
            inject = new if a <= pos < b else ""  # 只有「含匹配起点」的 run 承载 new
            node.text = prefix + inject + suffix
            if inject:
                node.set(XML_SPACE, "preserve")
        count += 1
        search_from = pos + len(new)


def replace_runs_with_token(
    docx_path: str | Path,
    replacements: Mapping[str, str],
    *,
    output: str | Path | None = None,
    strict: bool = True,
) -> int:
    """把 ``.docx`` 正文里的**具体实值**替换成 ``{{token}}`` 占位（跨 run 合并），返回替换总次数。

    Replace concrete values in a ``.docx`` with ``{{token}}`` placeholders (run-splitting aware).

    这是 W3 抽模板的「文本占位」方向：把参考文件里的实名/日期/金额等换成 docxtpl 的 ``{{jinja}}`` 占位，
    产出的模板可被 ``docxtpl`` 渲染回填。``replacements = {实值: 占位串}``（如 ``{"张三": "{{name}}"}``）。

    - **消解 run-splitting**：实值可能被 Word 拆进多个 run，本函数按段内拼接全文定位并把 ``{{token}}``
      落进单个 run——这正是手工敲占位常被自动更正拆断、导致 docxtpl 认不出的根因所在。
    - ``strict=True``（默认）时，任一 ``old`` 在全文一次都没命中即抛 :class:`AnchorNotFoundError`（fail-fast，
      与 :func:`fill_sdt_controls` 一致）；``strict=False`` 则静默跳过未命中项。
    - 仅处理正文 ``word/document.xml``；``output=None`` 原地写回。同一段落按 ``replacements`` 的插入顺序
      依次施用。
    """
    infos, items = read_package(docx_path)
    doc = parse_xml(items["word/document.xml"])
    body = doc.find(qn("w:body"))
    matched = dict.fromkeys(replacements, 0)
    total = 0
    if body is not None:
        for para in body.iter(qn("w:p")):
            t_nodes = list(para.iter(qn("w:t")))
            if not t_nodes:
                continue
            for old, new in replacements.items():
                hits = _replace_across_runs(t_nodes, old, new)
                matched[old] += hits
                total += hits

    if strict:
        missing = [old for old, hits in matched.items() if hits == 0]
        if missing:
            raise AnchorNotFoundError(
                f"这些实值未在文档中找到 | values not found in document: {missing}",
            )

    items["word/document.xml"] = serialize_xml(doc)
    write_package(output or docx_path, infos, items)
    return total
