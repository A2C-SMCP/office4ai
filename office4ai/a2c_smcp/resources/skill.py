# filename: skill.py
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
"""``skill://`` Resource —— 通过 MCP Resource 分发 SKILL 能力包（A2C SKILL 通道 producer 侧）。

``skill://`` Resource — distribute a SKILL capability package over MCP Resources (A2C SKILL
channel, producer side).

协议依据 / Protocol: a2c-smcp-protocol ``docs/specification/skill.md`` §3（``resources`` source 模式）、
§11（MCP Server 实现指南）。office4ai 只做 **producer 侧零协议改动实现**：把本地 SKILL 文件夹
（``SKILL.md`` + ``scripts/`` + ``references/`` + 二进制模板资产）经 ``resources/list`` + ``resources/read``
暴露，Computer 负责物化（staging）到本地、合成全局 name ``mcp:<server>:<frontmatter.name>``。

**为何选 ``resources`` 模式**（milestone #4 · S3，见 ``docs/milestone4_authoring_desktop_spec.md``）：
- 传输无关——stdio / SSE / streamable-http 均可（``mounted`` 要求 producer 与 Computer 同机）；
- python-sdk 唯一端到端集成验证过的模式（``tests/.../fastmcp_skill_stdio_server.py``）；
- 天然承载 ``scripts/`` 分发：每个包内文件作为兄弟资源 ``skill://<host>/<leaf>/<相对路径>`` 暴露。

**契约要点**（由 Computer 消费方 ``staging.py`` 验证）：
- **SKILL 根**：``skill://<host>/<leaf>``，带 ``_meta.source="resources"``（缺则被当子资源、不注册为根）；
- **子资源**：同一 ``resources/list`` 内 URI = ``skill://<host>/<leaf>/<rel>`` 的兄弟资源，**不带** ``_meta``；
- 物化后包根须含 ``SKILL.md``，其 YAML frontmatter 至少有 ``name`` + ``description``。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from loguru import logger
from mcp.server.lowlevel.helper_types import ReadResourceContents
from mcp.types import Resource

from office4ai.a2c_smcp.resources.base import BaseResource

# SKILL URI 命名空间 host（反向域名，skill.md §11.2 推荐）/ SKILL URI namespace host (reverse-domain).
DEFAULT_SKILL_HOST = "com.a2c-smcp.office4ai"

SKILL_MD = "SKILL.md"

# 打包时忽略的开发噪音（不进 skill:// 分发）/ Dev noise excluded from the shipped package tree.
_IGNORED_NAMES = frozenset({".DS_Store", ".git", ".gitkeep", "__pycache__"})
_IGNORED_SUFFIXES = frozenset({".pyc", ".pyo"})

# 确定性 扩展名 → MIME 映射（producer 侧内容类型；文本子集经 UTF-8 内联，其余转 base64 blob）。
# Deterministic ext → MIME map (producer content type); the text subset is inlined as UTF-8,
# everything else is shipped as a base64 blob (BlobResourceContents).
_TEXT_MIME: dict[str, str] = {
    ".md": "text/markdown",
    ".markdown": "text/markdown",
    ".txt": "text/plain",
    ".py": "text/x-python",
    ".json": "application/json",
    ".yaml": "application/yaml",
    ".yml": "application/yaml",
    ".toml": "application/toml",
    ".xml": "application/xml",
    ".rst": "text/x-rst",
    ".csv": "text/csv",
    ".html": "text/html",
    ".css": "text/css",
    ".js": "application/javascript",
    ".ini": "text/plain",
    ".cfg": "text/plain",
    ".sh": "application/x-sh",
}
_BINARY_MIME: dict[str, str] = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".svg": "image/svg+xml",
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".potx": "application/vnd.openxmlformats-officedocument.presentationml.template",
}
_OCTET_STREAM = "application/octet-stream"


class SkillResourceError(Exception):
    """SKILL 包结构非法（缺 SKILL.md / frontmatter 缺字段 / name 非 kebab / 路径穿越）。"""


# 严格 kebab（marketplace SKILL v1 §3.1，消费方 naming lexer 同契约）：[a-z0-9-]，不以 - 始末、无连续 --。
_SKILL_NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


def _is_valid_skill_name(name: str) -> bool:
    """校验 SKILL name 为严格 kebab（1–64）。producer 侧提前拒收，避免广告一个消费方随后拒绝的 SKILL 根。"""
    return bool(name) and len(name) <= 64 and _SKILL_NAME_RE.match(name) is not None


def _guess_mime(name: str) -> str:
    ext = Path(name).suffix.lower()
    if ext in _TEXT_MIME:
        return _TEXT_MIME[ext]
    return _BINARY_MIME.get(ext, _OCTET_STREAM)


def _is_text(name: str) -> bool:
    return Path(name).suffix.lower() in _TEXT_MIME


def _is_ignored(rel_parts: tuple[str, ...]) -> bool:
    """任一路径段属噪音目录/文件名（``rel_parts[-1]`` 即文件名）或后缀属噪音 → 忽略。"""
    if any(part in _IGNORED_NAMES for part in rel_parts):
        return True
    return Path(rel_parts[-1]).suffix.lower() in _IGNORED_SUFFIXES


def parse_skill_frontmatter(text: str) -> dict[str, str]:
    """解析 SKILL.md 头部 ``---`` fenced frontmatter 的**标量**键值 / Parse scalar keys of SKILL.md frontmatter.

    仅取顶层 ``key: value`` 标量（``name`` / ``description`` / ``version`` 足矣），不引入 YAML 依赖：
    office4ai 自有 SKILL 的 frontmatter 由我们编写、格式可控。无 frontmatter / 无闭合 ``---`` → ``{}``。
    Only top-level scalar ``key: value`` pairs are extracted (no YAML dependency); office4ai authors
    its own SKILL frontmatter so the format is controlled. No / unclosed fence → ``{}``.
    """
    if not text.startswith("---"):
        return {}
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return {}
    fm: dict[str, str] = {}
    for line in lines[1:]:
        if line.strip() == "---":
            return fm
        key, sep, raw = line.partition(":")
        key = key.strip()
        if not sep or not key:
            continue
        val = raw.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
            val = val[1:-1]
        fm[key] = val
    # 无闭合 fence → 视为无有效 frontmatter（与消费方 staging.py 语义一致）。
    return {}


class SkillResource(BaseResource):
    """把一个本地 SKILL 文件夹经 ``resources`` source 模式暴露为 ``skill://`` 资源。

    Expose one local SKILL folder as a ``skill://`` resource in ``resources`` source mode.

    :param skill_dir: SKILL 包根目录（须含 ``SKILL.md``）。
    :param host: URI 命名空间 host（反向域名）；同一 server 内 URI 唯一。
    """

    def __init__(self, skill_dir: Path, *, host: str = DEFAULT_SKILL_HOST) -> None:
        self._dir = Path(skill_dir).resolve()
        self._host = host

        skill_md = self._dir / SKILL_MD
        if not skill_md.is_file():
            raise SkillResourceError(f"SKILL package missing {SKILL_MD}: {self._dir}")

        fm = parse_skill_frontmatter(skill_md.read_text(encoding="utf-8"))
        name = fm.get("name")
        description = fm.get("description")
        if not name or not description:
            raise SkillResourceError(f"{SKILL_MD} frontmatter must define non-empty 'name' + 'description': {skill_md}")
        # name 须与消费方 lexer 同契约（严格 kebab），否则 Computer 会拒绝注册该 SKILL —— 提前拒收保持一致。
        if not _is_valid_skill_name(name):
            raise SkillResourceError(
                f"{SKILL_MD} frontmatter 'name' must be strict kebab (§3.1), got {name!r}: {skill_md}"
            )

        # 约定：frontmatter.name == 目录 basename（消费方 staging 会以 frontmatter.name 校正包根目录名）。
        self._name = name
        self._description = description
        # version 非 frontmatter 强制字段（skill.md §6）；若 SKILL 声明则透传进 _meta.version（producer 便利）。
        self._version: str | None = fm.get("version") or None

    # ── BaseResource 身份 / identity ──────────────────────────────────────────

    @property
    def base_uri(self) -> str:
        return f"skill://{self._host}/{self._name}"

    @property
    def uri(self) -> str:
        # 根资源无 query；子资源 URI 在 list_entries 内独立合成。
        return self.base_uri

    @property
    def name(self) -> str:
        return self._name

    @property
    def description(self) -> str:
        return self._description

    @property
    def mime_type(self) -> str:
        # 根 = 目录包（与 python-sdk 参考 producer fixture 一致）。
        return "inode/directory"

    @property
    def meta(self) -> dict[str, Any] | None:
        m: dict[str, Any] = {"source": "resources"}
        if self._version is not None:
            m["version"] = self._version
        return m

    # ── resources/list 展开：根（带 _meta）+ 每个包内文件一个子资源（无 _meta）──────────

    def list_entries(self) -> list[Resource]:
        entries: list[Resource] = [
            Resource.model_validate(
                {
                    "uri": self.base_uri,
                    "name": self._name,
                    "description": self._description,
                    "mimeType": self.mime_type,
                    "_meta": self.meta,
                },
            ),
        ]
        for rel in self._iter_rel_files():
            entries.append(
                Resource.model_validate(
                    {
                        "uri": f"{self.base_uri}/{rel}",
                        "name": rel,
                        "mimeType": _guess_mime(rel),
                    },
                ),
            )
        return entries

    # ── resources/read：按相对路径服务文本 / 二进制内容 ─────────────────────────────

    async def read(self) -> str:
        """根读取 → 返回 SKILL.md 原文（满足抽象接口；消费方实际逐个读子资源）。"""
        return (self._dir / SKILL_MD).read_text(encoding="utf-8")

    async def read_content(self, uri: str) -> list[ReadResourceContents]:
        rel = self._rel_from_uri(uri)
        target = self._safe_join(rel) if rel else (self._dir / SKILL_MD)
        data = target.read_bytes()
        mime = _guess_mime(target.name)
        if _is_text(target.name):
            # 扩展名判文本仅是乐观假设：内容非 UTF-8 时回退为二进制 blob，避免抛错拖垮整个 SKILL 物化。
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                logger.warning(f"text-extension SKILL file is not UTF-8, shipping as blob: {target}")
                return [ReadResourceContents(content=data, mime_type=mime)]
            return [ReadResourceContents(content=text, mime_type=mime)]
        return [ReadResourceContents(content=data, mime_type=mime)]

    # ── 内部 / internal ────────────────────────────────────────────────────────

    def _within_package(self, path: Path) -> bool:
        """path 的 realpath 是否仍在包根内（拒绝符号链接逃逸）。"""
        resolved = path.resolve()
        return resolved == self._dir or self._dir in resolved.parents

    def _iter_rel_files(self) -> list[str]:
        """包内所有文件的 POSIX 相对路径（确定序，跳过开发噪音与逃出包根的符号链接）。

        list 阶段与 read 阶段（``_safe_join`` 的 realpath 校验）保持**同一可达集**：符号链接逃出
        包根者不广告，避免出现「已列出但读必失败」的子资源。
        """
        rels: list[str] = []
        for path in sorted(self._dir.rglob("*")):
            if not path.is_file():
                continue
            rel = path.relative_to(self._dir)
            if _is_ignored(rel.parts):
                continue
            if not self._within_package(path):
                logger.warning(f"SKILL file escapes package via symlink, skipped: {path}")
                continue
            rels.append(rel.as_posix())
        return rels

    def _rel_from_uri(self, uri: str) -> str:
        """从读 URI 提取包内相对路径；根 URI → ``""``。"""
        if uri == self.base_uri:
            return ""
        prefix = self.base_uri + "/"
        if uri.startswith(prefix):
            return uri[len(prefix) :]
        raise SkillResourceError(f"URI not owned by this SKILL ({self.base_uri!r}): {uri!r}")

    def _safe_join(self, rel: str) -> Path:
        """安全拼接包内相对路径，防 ``..`` / 绝对路径穿越，并校验文件存在。"""
        target = (self._dir / rel).resolve()
        if target != self._dir and self._dir not in target.parents:
            raise SkillResourceError(f"path escapes SKILL package: {rel!r}")
        if not target.is_file():
            raise SkillResourceError(f"SKILL resource not found: {rel!r}")
        return target


def discover_skill_resources(skills_root: Path | str, *, host: str = DEFAULT_SKILL_HOST) -> list[SkillResource]:
    """扫描 ``skills_root`` 下**一级**子目录，凡含 ``SKILL.md`` 者建一个 ``SkillResource``。

    Scan first-level subdirs of ``skills_root``; each subdir containing ``SKILL.md`` becomes a
    ``SkillResource``. 单个 SKILL 解析失败（缺 SKILL.md / frontmatter 缺字段 / name 非 kebab /
    SKILL.md 不可读或非 UTF-8）记 ERROR 跳过、**不阻断**其余——对齐消费方 staging 的部分失败隔离
    （skill.md §1.5：batch 接口对部分失败健壮），一个坏包不得 brick 整个 server 启动。root 不存在 / 为空 → ``[]``。
    """
    root = Path(skills_root)
    if not root.is_dir():
        return []
    out: list[SkillResource] = []
    for child in sorted(root.iterdir()):
        if not child.is_dir() or not (child / SKILL_MD).is_file():
            continue
        try:
            out.append(SkillResource(child, host=host))
        except Exception as e:  # 坏包隔离：SkillResourceError / OSError / UnicodeDecodeError 等一律跳过、不阻断其余
            logger.error(f"Skipping invalid SKILL package {child}: {e}")
    return out
