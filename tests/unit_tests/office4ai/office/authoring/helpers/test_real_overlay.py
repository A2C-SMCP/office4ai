"""真实微软模板叠加测试（本地可选）| Real Microsoft-template overlay tests (opt-in).

仅当设置 ``OFFICE4AI_REAL_TEMPLATE_DIR`` 且 ``soffice`` 可用时运行；CI 上默认 skip。
验证原语对**真实 Office 模板**（而非自制合成件）可用——这是 S2 验收「对三平台真实模板可用」的凭据。
叠加目录须含 ``real_word.dotx`` / ``real_ppt.potx`` / ``real_excel.xltx``（不提交进仓）。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from pptx import Presentation

from office4ai.office.authoring.helpers import (
    fill_sdt_controls,
    instantiate_from_template,
    locate_by_named_range,
)
from office4ai.office.authoring.helpers._ooxml import parse_xml, qn, read_package


def _count_text_sdts(docx: Path) -> int:
    _, items = read_package(docx)
    doc = parse_xml(items["word/document.xml"])
    count = 0
    for sdt in doc.iter(qn("w:sdt")):
        content = sdt.find(qn("w:sdtContent"))
        if content is None:
            continue
        texts = list(content.iter(qn("w:t")))
        if texts and "".join(t.text or "" for t in texts).strip():
            count += 1
    return count


def test_real_dotx_instantiate_and_fill(real_template_dir: Path, tmp_path: Path, soffice_or_skip: str) -> None:
    src = real_template_dir / "real_word.dotx"
    if not src.exists():
        pytest.skip("叠加目录缺 real_word.dotx")
    doc_path = instantiate_from_template(src, tmp_path / "real.docx")
    Document(doc_path)  # 合法文档
    n = _count_text_sdts(doc_path)
    assert n > 0, "真实模板应含文字 SDT 控件"
    # 用等量占位值做非严格填充，验证 fill_sdt_controls 对真实模板可用
    filled = fill_sdt_controls(doc_path, [f"值{i}" for i in range(n)], strict=False)
    assert filled == n
    Document(doc_path)  # 填充后仍是合法文档


def test_real_potx_instantiate(real_template_dir: Path, tmp_path: Path) -> None:
    src = real_template_dir / "real_ppt.potx"
    if not src.exists():
        pytest.skip("叠加目录缺 real_ppt.potx")
    out = instantiate_from_template(src, tmp_path / "real.pptx")
    prs = Presentation(out)  # 合法演示文稿
    assert len(prs.slide_masters[0].slide_layouts) > 0


def test_real_xltx_instantiate_and_locate(real_template_dir: Path, tmp_path: Path, soffice_or_skip: str) -> None:
    src = real_template_dir / "real_excel.xltx"
    if not src.exists():
        pytest.skip("叠加目录缺 real_excel.xltx")
    out = instantiate_from_template(src, tmp_path / "real.xlsx")
    assert out.exists()
    # 至少定位到模板里已知的一个命名区域（模板自带定义名）
    found_any = any(locate_by_named_range(src, name) for name in ("选定年份", "年份列表", "指标列表"))
    assert found_any, "真实 .xltx 应能定位到其定义名"
