"""SkillResource 单元测试 | SkillResource unit tests (milestone #4 · S3, issue #59).

覆盖 skill:// producer 侧 resources source 模式：身份 / list_entries 展开（根带 _meta、子资源不带）/
read_content 文本+二进制 / 路径穿越守卫 / frontmatter 解析 / 目录发现。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from office4ai.a2c_smcp.resources.skill import (
    DEFAULT_SKILL_HOST,
    SkillResource,
    SkillResourceError,
    discover_skill_resources,
    parse_skill_frontmatter,
)

FIXTURE_SKILLS = Path(__file__).parent / "fixtures" / "skills"
DEMO_DIR = FIXTURE_SKILLS / "demo-authoring"
DEMO_BASE = f"skill://{DEFAULT_SKILL_HOST}/demo-authoring"


@pytest.fixture
def demo() -> SkillResource:
    return SkillResource(DEMO_DIR)


# ── 身份 / identity ─────────────────────────────────────────────────────────
class TestSkillResourceIdentity:
    def test_base_uri(self, demo: SkillResource) -> None:
        assert demo.base_uri == DEMO_BASE

    def test_uri_equals_base(self, demo: SkillResource) -> None:
        assert demo.uri == demo.base_uri

    def test_name(self, demo: SkillResource) -> None:
        assert demo.name == "demo-authoring"

    def test_description(self, demo: SkillResource) -> None:
        assert demo.description.startswith("Demo authoring SKILL fixture")

    def test_mime_type_is_directory(self, demo: SkillResource) -> None:
        assert demo.mime_type == "inode/directory"

    def test_meta_source_and_version(self, demo: SkillResource) -> None:
        assert demo.meta == {"source": "resources", "version": "0.1.0"}

    def test_custom_host(self) -> None:
        res = SkillResource(DEMO_DIR, host="com.example.skills")
        assert res.base_uri == "skill://com.example.skills/demo-authoring"


# ── list_entries：根 + 子资源 / registrable root + sub-resources ──────────────
class TestSkillResourceListEntries:
    def test_root_carries_meta_source(self, demo: SkillResource) -> None:
        entries = demo.list_entries()
        root = entries[0]
        assert str(root.uri) == DEMO_BASE
        assert root.meta is not None
        assert root.meta.get("source") == "resources"
        assert root.meta.get("version") == "0.1.0"
        assert root.mimeType == "inode/directory"

    def test_sub_resources_present_and_meta_absent(self, demo: SkillResource) -> None:
        entries = demo.list_entries()
        sub_by_uri = {str(r.uri): r for r in entries[1:]}
        expected = {
            f"{DEMO_BASE}/SKILL.md",
            f"{DEMO_BASE}/scripts/gen_demo.py",
            f"{DEMO_BASE}/references/notes.md",
            f"{DEMO_BASE}/assets/logo.png",
        }
        assert expected <= set(sub_by_uri)
        # 子资源一律不带 _meta（否则会被消费方当作独立 SKILL 根）
        for uri in expected:
            assert sub_by_uri[uri].meta is None

    def test_sub_resource_mime_types(self, demo: SkillResource) -> None:
        by_uri = {str(r.uri): r for r in demo.list_entries()[1:]}
        assert by_uri[f"{DEMO_BASE}/SKILL.md"].mimeType == "text/markdown"
        assert by_uri[f"{DEMO_BASE}/scripts/gen_demo.py"].mimeType == "text/x-python"
        assert by_uri[f"{DEMO_BASE}/assets/logo.png"].mimeType == "image/png"

    def test_ignores_dev_noise(self, tmp_path: Path) -> None:
        skill = tmp_path / "noisy"
        (skill / "scripts" / "__pycache__").mkdir(parents=True)
        (skill / "SKILL.md").write_text("---\nname: noisy\ndescription: has noise\n---\nbody\n", encoding="utf-8")
        (skill / "scripts" / "helper.py").write_text("x = 1\n", encoding="utf-8")
        (skill / "scripts" / "__pycache__" / "helper.cpython-311.pyc").write_bytes(b"\x00\x01")
        (skill / ".DS_Store").write_bytes(b"\x00")
        res = SkillResource(skill)
        rels = {str(r.uri).removeprefix(res.base_uri + "/") for r in res.list_entries()[1:]}
        assert rels == {"SKILL.md", "scripts/helper.py"}

    def test_excludes_symlink_escaping_package(self, tmp_path: Path) -> None:
        # list 与 read 同一可达集：逃出包根的符号链接不广告（否则「已列出但读必失败」）
        outside = tmp_path / "outside.txt"
        outside.write_text("secret", encoding="utf-8")
        skill = tmp_path / "s"
        skill.mkdir()
        (skill / "SKILL.md").write_text("---\nname: s\ndescription: d\n---\n", encoding="utf-8")
        try:
            (skill / "leak.txt").symlink_to(outside)
        except (OSError, NotImplementedError):
            pytest.skip("symlink unsupported on this platform")
        res = SkillResource(skill)
        rels = {str(r.uri).removeprefix(res.base_uri + "/") for r in res.list_entries()[1:]}
        assert rels == {"SKILL.md"}


# ── read_content：文本 / 二进制 / 守卫 ────────────────────────────────────────
class TestSkillResourceReadContent:
    @pytest.mark.asyncio
    async def test_read_root_returns_skill_md(self, demo: SkillResource) -> None:
        contents = await demo.read_content(DEMO_BASE)
        assert len(contents) == 1
        assert isinstance(contents[0].content, str)
        assert "name: demo-authoring" in contents[0].content
        assert contents[0].mime_type == "text/markdown"

    @pytest.mark.asyncio
    async def test_read_text_subfile(self, demo: SkillResource) -> None:
        contents = await demo.read_content(f"{DEMO_BASE}/scripts/gen_demo.py")
        assert isinstance(contents[0].content, str)
        assert "def main()" in contents[0].content
        assert contents[0].mime_type == "text/x-python"

    @pytest.mark.asyncio
    async def test_read_binary_subfile_returns_bytes(self, demo: SkillResource) -> None:
        contents = await demo.read_content(f"{DEMO_BASE}/assets/logo.png")
        assert isinstance(contents[0].content, bytes)
        assert contents[0].content[:8] == b"\x89PNG\r\n\x1a\n"  # PNG magic
        assert contents[0].mime_type == "image/png"

    @pytest.mark.asyncio
    async def test_read_matches_disk(self, demo: SkillResource) -> None:
        contents = await demo.read_content(f"{DEMO_BASE}/references/notes.md")
        assert contents[0].content == (DEMO_DIR / "references" / "notes.md").read_text(encoding="utf-8")

    @pytest.mark.asyncio
    async def test_traversal_rejected(self, demo: SkillResource) -> None:
        with pytest.raises(SkillResourceError, match="escapes SKILL package"):
            await demo.read_content(f"{DEMO_BASE}/../../secret.txt")

    @pytest.mark.asyncio
    async def test_missing_subfile(self, demo: SkillResource) -> None:
        with pytest.raises(SkillResourceError, match="not found"):
            await demo.read_content(f"{DEMO_BASE}/nope.md")

    @pytest.mark.asyncio
    async def test_unowned_uri_rejected(self, demo: SkillResource) -> None:
        with pytest.raises(SkillResourceError, match="not owned"):
            await demo.read_content("skill://other.host/other/SKILL.md")

    @pytest.mark.asyncio
    async def test_read_returns_skill_md_body(self, demo: SkillResource) -> None:
        # 抽象接口 read()：根读取回 SKILL.md 原文（含 frontmatter）
        text = await demo.read()
        assert text == (DEMO_DIR / "SKILL.md").read_text(encoding="utf-8")

    @pytest.mark.asyncio
    async def test_text_ext_non_utf8_falls_back_to_blob(self, tmp_path: Path) -> None:
        # 扩展名判文本但内容非 UTF-8 → 回退二进制 blob，不抛错
        skill = tmp_path / "s"
        skill.mkdir()
        (skill / "SKILL.md").write_text("---\nname: s\ndescription: d\n---\n", encoding="utf-8")
        (skill / "data.csv").write_bytes(b"\xff\xfe\x00 not utf-8")
        res = SkillResource(skill)
        contents = await res.read_content(f"{res.base_uri}/data.csv")
        assert isinstance(contents[0].content, bytes)
        assert contents[0].content == b"\xff\xfe\x00 not utf-8"


# ── 构造校验 / construction validation ────────────────────────────────────────
class TestSkillResourceConstruction:
    def test_missing_skill_md(self, tmp_path: Path) -> None:
        (tmp_path / "empty").mkdir()
        with pytest.raises(SkillResourceError, match="missing SKILL.md"):
            SkillResource(tmp_path / "empty")

    def test_frontmatter_missing_name(self, tmp_path: Path) -> None:
        skill = tmp_path / "s"
        skill.mkdir()
        (skill / "SKILL.md").write_text("---\ndescription: no name here\n---\nbody\n", encoding="utf-8")
        with pytest.raises(SkillResourceError, match="name.*description"):
            SkillResource(skill)

    def test_frontmatter_missing_description(self, tmp_path: Path) -> None:
        skill = tmp_path / "s"
        skill.mkdir()
        (skill / "SKILL.md").write_text("---\nname: s\n---\nbody\n", encoding="utf-8")
        with pytest.raises(SkillResourceError):
            SkillResource(skill)

    @pytest.mark.parametrize("bad_name", ["Bad_Name", "-lead", "trail-", "a--b", "UPPER", "has space"])
    def test_frontmatter_name_not_kebab_rejected(self, tmp_path: Path, bad_name: str) -> None:
        # producer 提前拒收非严格 kebab 的 name（否则消费方 lexer 会拒绝注册，形成广告/拒收错位）
        skill = tmp_path / "s"
        skill.mkdir()
        (skill / "SKILL.md").write_text(f'---\nname: "{bad_name}"\ndescription: d\n---\n', encoding="utf-8")
        with pytest.raises(SkillResourceError, match="kebab"):
            SkillResource(skill)


# ── frontmatter 解析 / frontmatter parsing ────────────────────────────────────
class TestFrontmatterParser:
    def test_basic(self) -> None:
        fm = parse_skill_frontmatter("---\nname: foo\ndescription: bar baz\n---\nbody\n")
        assert fm == {"name": "foo", "description": "bar baz"}

    def test_quoted_values(self) -> None:
        fm = parse_skill_frontmatter("---\nname: \"foo\"\ndescription: 'bar'\n---\n")
        assert fm["name"] == "foo"
        assert fm["description"] == "bar"

    def test_colon_in_value(self) -> None:
        fm = parse_skill_frontmatter("---\ndescription: a: b: c\n---\n")
        assert fm["description"] == "a: b: c"

    def test_no_fence(self) -> None:
        assert parse_skill_frontmatter("# just a heading\n") == {}

    def test_dashes_prefix_but_not_fence(self) -> None:
        # 以 "---" 开头但首行非纯 "---"（如 "---foo"）→ 无有效 frontmatter
        assert parse_skill_frontmatter("---foo\nname: x\n") == {}

    def test_unclosed_fence(self) -> None:
        assert parse_skill_frontmatter("---\nname: foo\nno closing fence\n") == {}


# ── 目录发现 / discovery ──────────────────────────────────────────────────────
class TestDiscovery:
    def test_discovers_demo(self) -> None:
        found = discover_skill_resources(FIXTURE_SKILLS)
        assert [r.name for r in found] == ["demo-authoring"]

    def test_nonexistent_root(self, tmp_path: Path) -> None:
        assert discover_skill_resources(tmp_path / "nope") == []

    def test_skips_dir_without_skill_md(self, tmp_path: Path) -> None:
        (tmp_path / "not-a-skill").mkdir()
        (tmp_path / "not-a-skill" / "readme.txt").write_text("hi", encoding="utf-8")
        assert discover_skill_resources(tmp_path) == []

    def test_keeps_valid_skips_invalid(self, tmp_path: Path) -> None:
        good = tmp_path / "good"
        good.mkdir()
        (good / "SKILL.md").write_text("---\nname: good\ndescription: ok\n---\n", encoding="utf-8")
        bad = tmp_path / "bad"
        bad.mkdir()
        (bad / "SKILL.md").write_text("---\nname: bad\n---\n", encoding="utf-8")  # missing description
        found = discover_skill_resources(tmp_path)
        assert [r.name for r in found] == ["good"]

    def test_non_utf8_skill_md_does_not_brick_discovery(self, tmp_path: Path) -> None:
        # 坏包隔离：一个 SKILL.md 非 UTF-8 不得抛出、拖垮其余（对齐消费方部分失败健壮性）
        good = tmp_path / "good"
        good.mkdir()
        (good / "SKILL.md").write_text("---\nname: good\ndescription: ok\n---\n", encoding="utf-8")
        broken = tmp_path / "aaa-broken"  # 排序在 good 之前，验证坏包不阻断后续
        broken.mkdir()
        (broken / "SKILL.md").write_bytes(b"\xff\xfe\x00 not utf-8")
        found = discover_skill_resources(tmp_path)
        assert [r.name for r in found] == ["good"]
