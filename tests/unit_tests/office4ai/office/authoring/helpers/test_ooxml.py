"""共享 OOXML 工具单测 | Unit tests for the shared OOXML helpers (_ooxml)."""

from __future__ import annotations

from pathlib import Path

import pytest

from office4ai.office.authoring.helpers._ooxml import (
    NS,
    parse_xml,
    qn,
    read_package,
    serialize_xml,
    write_package,
)


class TestQn:
    def test_known_prefix_expands_to_clark(self) -> None:
        assert qn("w:sdt") == f"{{{NS['w']}}}sdt"

    def test_bare_tag_without_prefix(self) -> None:
        assert qn("Relationship") == "Relationship"

    def test_unknown_prefix_raises_keyerror(self) -> None:
        with pytest.raises(KeyError):
            qn("zzz:tag")


class TestPackageRoundTrip:
    def test_read_modify_write_preserves_other_parts(self, chart_xlsx: Path, tmp_path: Path) -> None:
        infos, items = read_package(chart_xlsx)
        original_names = {i.filename for i in infos}
        # 改一个部件，回写，确认其余部件与部件清单不变
        root = parse_xml(items["xl/workbook.xml"])
        items["xl/workbook.xml"] = serialize_xml(root)
        out = tmp_path / "rt.xlsx"
        write_package(out, infos, items)
        infos2, _ = read_package(out)
        assert {i.filename for i in infos2} == original_names
