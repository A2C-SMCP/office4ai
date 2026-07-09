"""fill_sdt_controls / locate_by_style 单测 | Unit tests for the Word primitives."""

from __future__ import annotations

from pathlib import Path

import pytest

from office4ai.office.authoring.helpers import (
    AnchorNotFoundError,
    SdtCountMismatchError,
    fill_sdt_controls,
    locate_by_style,
)
from office4ai.office.authoring.helpers._ooxml import parse_xml, qn, read_package


def _document_xml(docx: Path) -> str:
    _, items = read_package(docx)
    return items["word/document.xml"].decode("utf-8")


def _sdt_texts(docx: Path) -> list[str]:
    """按文档顺序取每个 SDT 的可见文本，用于断言填充落点。"""
    _, items = read_package(docx)
    doc = parse_xml(items["word/document.xml"])
    out: list[str] = []
    for sdt in doc.iter(qn("w:sdt")):
        content = sdt.find(qn("w:sdtContent"))
        if content is None:
            continue
        texts = list(content.iter(qn("w:t")))
        if texts:
            out.append("".join(t.text or "" for t in texts))
    return out


class TestFillSdtControls:
    def test_list_mode_fills_in_order(self, sdt_docx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        n = fill_sdt_controls(sdt_docx, ["标题A", "副标题B", "章节C"], output=out)
        assert n == 3
        assert _sdt_texts(out) == ["标题A", "副标题B", "章节C"]

    def test_dict_mode_matches_by_alias(self, sdt_docx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        n = fill_sdt_controls(
            sdt_docx,
            {"文档标题": "2026 报告", "副标题": "副标题内容", "章节一标题": "第一章"},
            output=out,
        )
        assert n == 3
        assert _sdt_texts(out) == ["2026 报告", "副标题内容", "第一章"]

    def test_removes_placeholder_marks(self, sdt_docx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        fill_sdt_controls(sdt_docx, ["a", "b", "c"], output=out)
        xml = _document_xml(out)
        assert "showingPlcHdr" not in xml
        assert 'xml:space="preserve"' in xml  # 首尾空格保护标记已加

    def test_in_place_edit(self, sdt_docx: Path) -> None:
        n = fill_sdt_controls(sdt_docx, ["x", "y", "z"])  # output=None -> 原地
        assert n == 3
        assert _sdt_texts(sdt_docx) == ["x", "y", "z"]

    def test_strict_count_mismatch_raises(self, sdt_docx: Path) -> None:
        with pytest.raises(SdtCountMismatchError):
            fill_sdt_controls(sdt_docx, ["only-one"])

    def test_non_strict_count_mismatch_allowed(self, sdt_docx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        n = fill_sdt_controls(sdt_docx, ["one", "two"], output=out, strict=False)
        assert n == 2  # 只填了前两个

    def test_dict_unknown_alias_strict_raises(self, sdt_docx: Path) -> None:
        with pytest.raises(AnchorNotFoundError):
            fill_sdt_controls(sdt_docx, {"不存在的alias": "v"})

    def test_dict_unknown_alias_non_strict_ok(self, sdt_docx: Path, tmp_path: Path) -> None:
        out = tmp_path / "out.docx"
        n = fill_sdt_controls(sdt_docx, {"文档标题": "T", "缺失": "x"}, output=out, strict=False)
        assert n == 1

    def test_rejects_str_values(self, sdt_docx: Path) -> None:
        with pytest.raises(TypeError):
            fill_sdt_controls(sdt_docx, "整串会被逐字符拆开")  # type: ignore[arg-type]


class TestLocateByStyle:
    def test_finds_heading_paragraphs(self, sdt_docx: Path) -> None:
        matches = locate_by_style(sdt_docx, "Heading1")
        texts = [m.text for m in matches]
        assert "普通章节正文" in texts
        assert all(isinstance(m.index, int) for m in matches)

    def test_unknown_style_returns_empty(self, sdt_docx: Path) -> None:
        assert locate_by_style(sdt_docx, "不存在的样式") == []
