"""create-office-file SKILL 包结构单测 | create-office-file SKILL package validity (milestone #4 · S4).

保证生产 SKILL 包被 SkillResource 正确识别（frontmatter 合法、结构完整），否则 server 启动扫描
skills_root 时会拒收该包。沙箱执行行为在集成测试覆盖。
"""

from __future__ import annotations

from pathlib import Path

import office4ai
from office4ai.a2c_smcp.resources.skill import SkillResource

SKILL_DIR = Path(office4ai.__file__).resolve().parent / "office" / "skills" / "create-office-file"


def test_skill_dir_exists() -> None:
    assert (SKILL_DIR / "SKILL.md").is_file()


def test_frontmatter_identity() -> None:
    res = SkillResource(SKILL_DIR)
    assert res.name == "create-office-file"  # 目录名 == frontmatter.name（严格 kebab）
    assert res.description  # 非空
    assert "skill://" in res.base_uri


def test_reference_scripts_and_docs_exposed() -> None:
    res = SkillResource(SKILL_DIR)
    rels = {str(r.uri).removeprefix(res.base_uri + "/") for r in res.list_entries()[1:]}
    expected = {
        "SKILL.md",
        "scripts/gen_word.py",
        "scripts/gen_ppt.py",
        "scripts/gen_excel.py",
        "scripts/from_template.py",
        "references/template-reuse.md",
        "references/pitfalls.md",
    }
    assert expected <= rels
