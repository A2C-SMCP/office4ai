"""Excel 原语：保图表改单元格（删公式缓存逼重算）+ 定位命名区域。

Excel primitives: edit cells while preserving charts (strip formula cache to force
recalc), and locate defined names.

提炼自参考脚本 ``gen_excel.py``。核心是 :func:`fill_cells_lxml`——真实模板常带图表 /
条件格式 / 结构化表格，openpyxl 的 load→save 会损坏图表；因此只用 zip + lxml 改单元格 XML、
不碰图表部件，并删掉同 sheet 内所有公式的缓存值，逼 Excel/LibreOffice 打开时重算。
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

from ._ooxml import XML_SPACE, parse_xml, qn, read_package, serialize_xml, write_package
from .errors import AnchorNotFoundError


@dataclass(frozen=True)
class NamedRange:
    """一个定义名(命名区域) | A workbook defined name.

    ``refers_to`` 为原始公式（如 ``当前月份!$B$18:$C$30``）；``sheet`` / ``ref`` 为尽力解析出的
    工作表名与区域；``local_sheet_id`` 非 ``None`` 表示是表级作用域的定义名。
    """

    name: str
    refers_to: str
    sheet: str | None
    ref: str | None
    local_sheet_id: int | None


def _sheet_targets(items: dict[str, bytes]) -> dict[str, str]:
    """建立「工作表名 -> 包内 sheet XML 路径」映射（经 workbook.xml + rels 解析）。"""
    workbook = parse_xml(items["xl/workbook.xml"])
    rels = parse_xml(items["xl/_rels/workbook.xml.rels"])
    rid_to_target: dict[str, str] = {}
    for rel in rels.findall(qn("pr:Relationship")):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            # Target 相对 xl/ 目录；兼容绝对路径写法 | resolve relative to xl/
            rid_to_target[rel_id] = target[1:] if target.startswith("/") else f"xl/{target}"
    result: dict[str, str] = {}
    for sheet in workbook.iter(qn("s:sheet")):
        name = sheet.get("name")
        rid = sheet.get(qn("r:id"))
        if name and rid and rid in rid_to_target:
            result[name] = rid_to_target[rid]
    return result


def _strip_formula_cache(sheet_root: etree._Element) -> None:
    """删除所有公式单元格的缓存值 ``<v>``，强制打开时重算。"""
    for cell in sheet_root.iter(qn("s:c")):
        if cell.find(qn("s:f")) is not None:
            cached = cell.find(qn("s:v"))
            if cached is not None:
                cell.remove(cached)


def _force_full_recalc(items: dict[str, bytes]) -> None:
    """在 ``xl/workbook.xml`` 的 ``calcPr`` 上设 ``fullCalcOnLoad="1"``。

    删缓存对 LibreOffice 足够，但个别 Excel 版本还需此标记才可靠地打开即全量重算
    （对齐参考脚本 openpyxl 路径的 ``fullCalcOnLoad=True``）。
    """
    key = "xl/workbook.xml"
    workbook = parse_xml(items[key])
    calc_pr = workbook.find(qn("s:calcPr"))
    if calc_pr is None:
        calc_pr = etree.Element(qn("s:calcPr"))
        # calcPr 须排在 definedNames 之后（无则 sheets 之后），遵循 CT_Workbook 元素顺序
        anchor = workbook.find(qn("s:definedNames"))
        if anchor is None:
            anchor = workbook.find(qn("s:sheets"))
        if anchor is not None:
            anchor.addnext(calc_pr)
        else:
            workbook.append(calc_pr)
    calc_pr.set("fullCalcOnLoad", "1")
    items[key] = serialize_xml(workbook)


def _set_cell(sheet_root: etree._Element, ref: str, value: object) -> bool:
    """把 value 写进 ``ref`` 单元格。支持 ``str`` (inlineStr) / ``bool`` (``t="b"``) /
    ``int``·``float`` (数字 ``<v>``) / ``None`` (清空)；其余类型抛 :class:`TypeError`。

    仅编辑**已存在**的 ``<c>`` 元素（模板占位单元格）；找不到返回 ``False``。
    """
    cell = sheet_root.find(f".//{qn('s:c')}[@r='{ref}']")
    if cell is None:
        return False
    cell.attrib.pop("t", None)
    for child in list(cell):
        cell.remove(child)
    if value is None:
        return True
    # bool 必须先于 int 判断（bool 是 int 子类），否则 True/False 会被当数字写成非法 <v>True</v>
    if isinstance(value, bool):
        cell.set("t", "b")
        etree.SubElement(cell, qn("s:v")).text = "1" if value else "0"
    elif isinstance(value, str):
        cell.set("t", "inlineStr")
        is_el = etree.SubElement(cell, qn("s:is"))
        text_el = etree.SubElement(is_el, qn("s:t"))
        text_el.text = value
        text_el.set(XML_SPACE, "preserve")
    elif isinstance(value, (int, float)):
        etree.SubElement(cell, qn("s:v")).text = str(value)
    else:
        raise TypeError(
            f"不支持的单元格值类型 {type(value).__name__} | unsupported cell value type"
            "（支持 str/bool/int/float/None；datetime 等请调用方先转为数字或字符串）",
        )
    return True


def fill_cells_lxml(
    template: str | Path,
    output: str | Path,
    sheet_cells: Mapping[str, Mapping[str, object]],
    *,
    strict: bool = False,
) -> int:
    """跨多个 sheet 批量改单元格，保图表/条件格式，并删公式缓存逼重算；返回成功写入的单元格数。

    Edit cells across sheets while preserving charts, then strip formula caches to force recalc.

    - ``sheet_cells = {工作表名: {单元格引用: 值或 None}}``；``None`` 清空该单元格。
    - 值类型：``str`` → inlineStr；``bool`` → ``t="b"`` (1/0)；``int``/``float`` → 数字 ``<v>``；
      其余类型抛 :class:`TypeError`（``datetime`` 等请调用方先转为数字或字符串）。
    - 每个涉及的 sheet 改完删其全部公式缓存，并在 workbook 上设 ``fullCalcOnLoad``，逼打开即重算。
    - 仅编辑模板中**已存在**的单元格；``strict=True`` 时若某引用不存在则抛
      :class:`AnchorNotFoundError`。
    - ``output`` 可与 ``template`` 相同（原地）；全程不经 openpyxl，避免真实模板的图表/扩展在
      round-trip 中丢失。公式的实际数值会在文件被 Excel/LibreOffice 打开时重算。
    """
    infos, items = read_package(template)
    targets = _sheet_targets(items)
    written = 0
    missing: list[str] = []

    for sheet_name, cells in sheet_cells.items():
        if sheet_name not in targets:
            raise AnchorNotFoundError(
                f"工作表未找到 | worksheet not found: {sheet_name!r}（可用 | available: {sorted(targets)}）",
            )
        target = targets[sheet_name]
        sheet_root = parse_xml(items[target])
        for ref, value in cells.items():
            if _set_cell(sheet_root, ref, value):
                written += 1
            else:
                missing.append(f"{sheet_name}!{ref}")
        _strip_formula_cache(sheet_root)
        items[target] = serialize_xml(sheet_root)

    if strict and missing:
        raise AnchorNotFoundError(f"这些单元格在模板中不存在 | cells not found in template: {missing}")

    _force_full_recalc(items)
    write_package(output, infos, items)
    return written


def _parse_refers_to(refers_to: str) -> tuple[str | None, str | None]:
    """把 ``Sheet!$A$1:$B$2`` 拆成 ``(sheet, ref)``；处理带引号的表名。"""
    if "!" not in refers_to:
        return None, refers_to or None
    sheet_part, _, ref = refers_to.rpartition("!")
    sheet = sheet_part.strip()
    if len(sheet) >= 2 and sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")  # OOXML 用 '' 转义单引号
    return (sheet or None), (ref or None)


def locate_by_named_range(xlsx_path: str | Path, name: str) -> list[NamedRange]:
    """定位工作簿中名为 ``name`` 的定义名(命名区域)，返回全部匹配（含表级作用域重名）。

    Locate workbook defined names matching ``name``. W3 抽模板时据此找出命名区域锚点。
    """
    _, items = read_package(xlsx_path)
    workbook = parse_xml(items["xl/workbook.xml"])
    matches: list[NamedRange] = []
    for defined in workbook.iter(qn("s:definedName")):
        if defined.get("name") != name:
            continue
        refers_to = (defined.text or "").strip()
        sheet, ref = _parse_refers_to(refers_to)
        local_raw = defined.get("localSheetId")
        matches.append(
            NamedRange(
                name=name,
                refers_to=refers_to,
                sheet=sheet,
                ref=ref,
                local_sheet_id=int(local_raw) if local_raw is not None else None,
            ),
        )
    return matches
