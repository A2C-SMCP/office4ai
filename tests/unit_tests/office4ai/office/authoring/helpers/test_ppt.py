"""create_custom_layout / locate_by_master 单测 | Unit tests for the PPT primitives."""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Emu

from office4ai.office.authoring.helpers import (
    AnchorNotFoundError,
    create_custom_layout,
    locate_by_master,
)

LAYOUT_NAME = "霓虹内容页"


def _decorate(shapes) -> None:
    """在临时页画一个带标记文字的文本框，便于断言装饰进入了版式。"""
    box = shapes.add_textbox(Emu(914400), Emu(914400), Emu(3000000), Emu(500000))
    box.text_frame.paragraphs[0].add_run().text = "BRAND_MARK"


class TestCreateCustomLayout:
    def test_registers_layout(self, blank_prs: Presentation) -> None:
        before = len(blank_prs.slide_masters[0].slide_layouts)
        layout = create_custom_layout(blank_prs, LAYOUT_NAME, _decorate)
        assert layout.name == LAYOUT_NAME
        assert len(blank_prs.slide_masters[0].slide_layouts) == before + 1

    def test_removes_temp_slide(self, blank_prs: Presentation) -> None:
        create_custom_layout(blank_prs, LAYOUT_NAME, _decorate)
        assert len(blank_prs.slides) == 0, "临时页应被删除，成品不应多一页"

    def test_decoration_lands_in_layout(self, blank_prs: Presentation, tmp_path: Path) -> None:
        create_custom_layout(blank_prs, LAYOUT_NAME, _decorate)
        out = tmp_path / "deck.pptx"
        blank_prs.save(out)
        layout = locate_by_master(Presentation(out), LAYOUT_NAME)
        text = "".join(
            run.text
            for shape in layout.shapes
            if shape.has_text_frame
            for para in shape.text_frame.paragraphs
            for run in para.runs
        )
        assert "BRAND_MARK" in text

    def test_survives_roundtrip_and_is_usable(self, blank_prs: Presentation, tmp_path: Path) -> None:
        create_custom_layout(blank_prs, LAYOUT_NAME, _decorate)
        out = tmp_path / "deck.pptx"
        blank_prs.save(out)
        reloaded = Presentation(out)
        layout = locate_by_master(reloaded, LAYOUT_NAME)
        reloaded.slides.add_slide(layout)  # 复用版式加一页，不应抛错
        used = tmp_path / "used.pptx"
        reloaded.save(used)
        assert len(Presentation(used).slides) == 1

    def test_bad_base_index_raises(self, blank_prs: Presentation) -> None:
        with pytest.raises(ValueError, match="out of range|越界"):
            create_custom_layout(blank_prs, LAYOUT_NAME, _decorate, base_layout_index=999)


class TestLocateByMaster:
    def test_found_after_create(self, blank_prs: Presentation) -> None:
        create_custom_layout(blank_prs, LAYOUT_NAME, _decorate)
        assert locate_by_master(blank_prs, LAYOUT_NAME).name == LAYOUT_NAME

    def test_not_found_raises(self, blank_prs: Presentation) -> None:
        with pytest.raises(AnchorNotFoundError):
            locate_by_master(blank_prs, "不存在的版式")
