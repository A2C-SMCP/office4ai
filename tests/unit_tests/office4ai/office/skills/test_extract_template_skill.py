"""extract-template SKILL 包结构单测 | extract-template SKILL package validity (milestone #4 · S6).

保证生产 SKILL 包被 SkillResource 正确识别；沙箱抽取行为在集成测试覆盖。
"""

from __future__ import annotations

from pathlib import Path

import office4ai
from office4ai.a2c_smcp.resources.skill import SkillResource

SKILL_DIR = Path(office4ai.__file__).resolve().parent / "office" / "skills" / "extract-template"


def test_skill_dir_exists() -> None:
    assert (SKILL_DIR / "SKILL.md").is_file()


def test_frontmatter_identity() -> None:
    res = SkillResource(SKILL_DIR)
    assert res.name == "extract-template"
    assert res.description
    assert "skill://" in res.base_uri


def test_reference_scripts_and_docs_exposed() -> None:
    res = SkillResource(SKILL_DIR)
    rels = {str(r.uri).removeprefix(res.base_uri + "/") for r in res.list_entries()[1:]}
    expected = {
        "SKILL.md",
        "scripts/extract_docx_sdt.py",
        "scripts/extract_docx_token.py",
        "scripts/extract_ppt_layout.py",
        "scripts/extract_xlsx_named.py",
        "references/anchors.md",
        "references/pitfalls.md",
    }
    assert expected <= rels
