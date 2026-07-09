"""authoring helper 单测的夹具 | Fixtures for the authoring helper unit tests.

committed 夹具全部**程序化生成真实 OOXML**（含 SDT 的 docx、含自定义母版/图表/公式的
pptx/xlsx，以及经 content-type 互换派生的自制 .dotx/.potx/.xltx）——自有 IP、确定性、CI 友好，
不提交任何微软专有模板。真实微软模板通过 ``OFFICE4AI_REAL_TEMPLATE_DIR`` 环境变量做本地可选叠加。
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn as docx_qn
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName
from pptx import Presentation
from pptx.util import Emu

from office4ai.office.authoring.helpers._ooxml import (
    CT_DOCX,
    CT_DOTX,
    CT_POTX,
    CT_PPTX,
    CT_XLSX,
    CT_XLTX,
    read_package,
    write_package,
)
from office4ai.office.authoring.helpers._soffice import SofficeNotFoundError, find_soffice

# ── SDT docx 构造 | build a docx with SDT content controls ──────────────────
#: 合成 SDT docx 的占位规格：(可见占位文本, alias, 段落样式或 None)。
SDT_SPECS = [
    ("【文档标题】", "文档标题", "Title"),
    ("【副标题】", "副标题", None),
    ("【章节一标题】", "章节一标题", "Heading 1"),
]


def _wrap_in_sdt(paragraph, alias: str, sdt_id: int) -> None:
    """把段落升级成 SDT 内容控件（占位符模式），复刻 gen_word.py 的构造手法。"""
    p = paragraph._p
    sdt = OxmlElement("w:sdt")
    sdt_pr = OxmlElement("w:sdtPr")
    for tag, val in (("w:alias", alias), ("w:tag", alias), ("w:id", str(sdt_id))):
        el = OxmlElement(tag)
        el.set(docx_qn("w:val"), val)
        sdt_pr.append(el)
    sdt_pr.append(OxmlElement("w:showingPlcHdr"))
    sdt.append(sdt_pr)
    content = OxmlElement("w:sdtContent")
    parent = p.getparent()
    idx = parent.index(p)
    parent.remove(p)
    content.append(p)
    sdt.append(content)
    parent.insert(idx, sdt)


@pytest.fixture
def sdt_docx(tmp_path: Path) -> Path:
    """含 3 个文字 SDT 控件（带 alias）+ 一个非 SDT 的 Heading1 段落的 .docx。"""
    doc = Document()
    for i, (text, alias, style) in enumerate(SDT_SPECS):
        para = doc.add_paragraph()
        para.add_run(text)
        if style:
            para.style = style
        _wrap_in_sdt(para, alias=alias, sdt_id=1000 + i)
    # 一个额外的、未包进 SDT 的 Heading1 段落，供 locate_by_style 命中
    extra = doc.add_paragraph("普通章节正文")
    extra.style = "Heading 1"
    path = tmp_path / "sdt.docx"
    doc.save(path)
    return path


# ── chart + formula xlsx 构造 | build an xlsx with a chart, formula, named range ──
@pytest.fixture
def chart_xlsx(tmp_path: Path) -> Path:
    """含图表 + 公式单元格 + 命名区域的 .xlsx（模拟真实仪表盘模板的关键特征）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    ws["A1"] = "项目"
    ws["B1"] = "金额"
    for i, (name, amount) in enumerate([("收入", 100), ("支出", 60)]):
        ws.cell(2 + i, 1, name)
        ws.cell(2 + i, 2, amount)
    ws["A5"] = "结余"
    ws["B5"] = "=B2-B3"  # 公式单元格：openpyxl 不算值，靠删缓存逼重算
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=3))
    ws.add_chart(chart, "D1")
    wb.defined_names["金额区"] = DefinedName("金额区", attr_text="数据!$B$2:$B$3")
    path = tmp_path / "chart.xlsx"
    wb.save(path)
    return path


# ── PPT 夹具 | presentation fixtures ────────────────────────────────────────
@pytest.fixture
def blank_prs() -> Presentation:
    """16:9 的空白 Presentation（含默认 7 个版式，index 6 = 空白），供 create_custom_layout。"""
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


# ── 自制「真实」模板：由文档经 content-type 互换派生 | self-authored templates ──
def _swap_content_type(src: Path, dst: Path, from_ct: str, to_ct: str) -> Path:
    infos, items = read_package(src)
    items["[Content_Types].xml"] = items["[Content_Types].xml"].replace(from_ct.encode(), to_ct.encode())
    write_package(dst, infos, items)
    return dst


@pytest.fixture
def potx_template(tmp_path: Path) -> Path:
    """自制 .potx：一份真实 pptx 经 presentation→template content-type 互换而来。"""
    pptx = tmp_path / "src.pptx"
    Presentation().save(pptx)
    return _swap_content_type(pptx, tmp_path / "brand.potx", CT_PPTX, CT_POTX)


@pytest.fixture
def dotx_template(sdt_docx: Path, tmp_path: Path) -> Path:
    """自制 .dotx：含 SDT 的真实 docx 经 document→template content-type 互换而来。"""
    return _swap_content_type(sdt_docx, tmp_path / "brand.dotx", CT_DOCX, CT_DOTX)


@pytest.fixture
def xltx_template(chart_xlsx: Path, tmp_path: Path) -> Path:
    """自制 .xltx：含图表/公式的真实 xlsx 经 sheet→template content-type 互换而来。"""
    return _swap_content_type(chart_xlsx, tmp_path / "brand.xltx", CT_XLSX, CT_XLTX)


# ── LibreOffice / 真实模板叠加 门控 | soffice & real-overlay gates ───────────
@pytest.fixture
def soffice_or_skip() -> str:
    """返回 soffice 路径；未安装则 skip 该测试（LibreOffice 实例化路径专用）。"""
    try:
        return find_soffice()
    except SofficeNotFoundError:
        pytest.skip("LibreOffice(soffice) 未安装，跳过实例化测试 | soffice not installed")


@pytest.fixture
def real_template_dir() -> Path:
    """真实微软模板叠加目录（``OFFICE4AI_REAL_TEMPLATE_DIR``）；未设置则 skip。"""
    raw = os.environ.get("OFFICE4AI_REAL_TEMPLATE_DIR")
    if not raw:
        pytest.skip("未设置 OFFICE4AI_REAL_TEMPLATE_DIR，跳过真实模板叠加测试 | real-template overlay unset")
    path = Path(raw)
    if not path.is_dir():
        pytest.skip(f"OFFICE4AI_REAL_TEMPLATE_DIR 不是目录 | not a directory: {path}")
    return path
