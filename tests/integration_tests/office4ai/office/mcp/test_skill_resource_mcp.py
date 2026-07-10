"""skill:// producer 侧 MCP 协议层集成测试 | skill:// producer-side MCP protocol integration tests.

milestone #4 · S3（issue #59）验收：
- ``resources/list`` 出现 ``skill://`` 根（带 ``_meta.source=resources``）+ 子资源（无 ``_meta``）；
- ``resources/read`` 服务文本（inline）与二进制（blob）子文件；
- 生产 server 经 ``OFFICE4AI_SKILLS_ROOT`` 暴露 SKILL 供 A2C Computer 物化（staging.py 消费的 wire 契约）。

测试 A 用 mcp 内存 client-server 会话驱动真 handlers（轻量、断言 _meta/blob 细节）；
测试 B 起真 OfficeMCPServer 子进程（证明生产二进制端到端暴露）。
"""

from __future__ import annotations

import base64
import os
from pathlib import Path

import pytest
from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client
from mcp.shared.memory import create_connected_server_and_client_session
from mcp.types import BlobResourceContents, TextResourceContents
from pydantic import AnyUrl

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.resources.skill import DEFAULT_SKILL_HOST, SkillResource
from office4ai.a2c_smcp.server import BaseMCPServer

# fixture SKILL 分发根（含 demo-authoring/）— 与单测共用。
FIXTURE_SKILLS = (
    Path(__file__).resolve().parents[4] / "unit_tests" / "office4ai" / "a2c_smcp" / "resources" / "fixtures" / "skills"
)
DEMO_BASE = f"skill://{DEFAULT_SKILL_HOST}/demo-authoring"
project_root = Path(__file__).resolve().parents[5]


class _SkillOnlyServer(BaseMCPServer):
    """仅注册 fixture SkillResource 的最小 server（无 workspace，用于内存协议层测试）。"""

    def _register_tools(self) -> None:
        pass

    def _register_resources(self) -> None:
        res = SkillResource(FIXTURE_SKILLS / "demo-authoring")
        self.resources[res.base_uri] = res


class _MultiSkillServer(BaseMCPServer):
    """注册多个 SkillResource（用于 _resolve_resource sibling-prefix 派发测试）。"""

    def __init__(self, config: MCPServerConfig, server_name: str, dirs: list[Path]) -> None:
        self._dirs = dirs
        super().__init__(config, server_name)

    def _register_tools(self) -> None:
        pass

    def _register_resources(self) -> None:
        for d in self._dirs:
            res = SkillResource(d)
            self.resources[res.base_uri] = res


def _write_skill(root: Path, name: str, marker: str) -> Path:
    d = root / name
    d.mkdir(parents=True)
    (d / "SKILL.md").write_text(
        f"---\nname: {name}\ndescription: {marker} desc\n---\n{marker} body\n", encoding="utf-8"
    )
    return d


@pytest.mark.integration
class TestSkillResourceProtocolInMemory:
    """内存 client-server 会话：断言 staging.py 消费的 wire 契约。"""

    async def test_list_exposes_skill_root_and_subresources(self) -> None:
        server = _SkillOnlyServer(MCPServerConfig(), "office4ai-skill-itest")
        async with create_connected_server_and_client_session(server.server) as session:
            result = await session.list_resources()
            by_uri = {str(r.uri): r for r in result.resources}

            # 根：带 _meta.source=resources（否则消费方不当作 SKILL 根）
            assert DEMO_BASE in by_uri
            root = by_uri[DEMO_BASE]
            assert root.meta is not None
            assert root.meta.get("source") == "resources"
            assert root.meta.get("version") == "0.1.0"

            # 子资源：URI 在 root+"/" 下，且**不带** _meta
            for rel in ("SKILL.md", "scripts/gen_demo.py", "references/notes.md", "assets/logo.png"):
                sub_uri = f"{DEMO_BASE}/{rel}"
                assert sub_uri in by_uri, f"missing sub-resource {sub_uri}"
                assert by_uri[sub_uri].meta is None

    async def test_read_text_subresource_inline(self) -> None:
        server = _SkillOnlyServer(MCPServerConfig(), "office4ai-skill-itest")
        async with create_connected_server_and_client_session(server.server) as session:
            result = await session.read_resource(AnyUrl(f"{DEMO_BASE}/SKILL.md"))
            assert len(result.contents) == 1
            content = result.contents[0]
            assert isinstance(content, TextResourceContents)
            assert "name: demo-authoring" in content.text
            assert content.mimeType == "text/markdown"

    async def test_read_binary_subresource_as_blob(self) -> None:
        server = _SkillOnlyServer(MCPServerConfig(), "office4ai-skill-itest")
        async with create_connected_server_and_client_session(server.server) as session:
            result = await session.read_resource(AnyUrl(f"{DEMO_BASE}/assets/logo.png"))
            content = result.contents[0]
            assert isinstance(content, BlobResourceContents)
            assert content.mimeType == "image/png"
            raw = base64.b64decode(content.blob)
            assert raw[:8] == b"\x89PNG\r\n\x1a\n"

    async def test_read_script_subresource(self) -> None:
        server = _SkillOnlyServer(MCPServerConfig(), "office4ai-skill-itest")
        async with create_connected_server_and_client_session(server.server) as session:
            result = await session.read_resource(AnyUrl(f"{DEMO_BASE}/scripts/gen_demo.py"))
            content = result.contents[0]
            assert isinstance(content, TextResourceContents)
            assert "def main()" in content.text


@pytest.mark.integration
class TestSkillResourceSiblingPrefixDispatch:
    """两个 leaf 互为前缀的 SKILL（demo / demo-authoring）：读取互不串台（_resolve_resource 守卫）。"""

    async def test_sibling_prefix_skills_do_not_crosswire(self, tmp_path: Path) -> None:
        _write_skill(tmp_path, "demo", "ALPHA")
        _write_skill(tmp_path, "demo-authoring", "BETA")
        server = _MultiSkillServer(
            MCPServerConfig(),
            "office4ai-multi-itest",
            [tmp_path / "demo", tmp_path / "demo-authoring"],
        )
        host = DEFAULT_SKILL_HOST
        async with create_connected_server_and_client_session(server.server) as session:
            a = await session.read_resource(AnyUrl(f"skill://{host}/demo/SKILL.md"))
            b = await session.read_resource(AnyUrl(f"skill://{host}/demo-authoring/SKILL.md"))
            a_text = a.contents[0]
            b_text = b.contents[0]
            assert isinstance(a_text, TextResourceContents) and isinstance(b_text, TextResourceContents)
            # demo 的读命中 demo（ALPHA），绝不串到 demo-authoring（BETA），反之亦然
            assert "ALPHA" in a_text.text and "BETA" not in a_text.text
            assert "BETA" in b_text.text and "ALPHA" not in b_text.text


@pytest.mark.integration
class TestSkillResourceRealServerSubprocess:
    """真 OfficeMCPServer 子进程（OFFICE4AI_SKILLS_ROOT 指向 fixture）：证明生产二进制端到端暴露。"""

    def _server_params(self) -> StdioServerParameters:
        return StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
            env={**os.environ, "OFFICE4AI_SKILLS_ROOT": str(FIXTURE_SKILLS)},
        )

    async def test_real_server_lists_skill_alongside_windows(self) -> None:
        async with stdio_client(self._server_params()) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.list_resources()
                uris = {str(r.uri) for r in result.resources}

                # 根索引 window 资源仍在（W4b-1：per-type 聚合窗口移除，无连接故无 per-file 窗口）
                assert any(u.startswith("window://office4ai?") for u in uris)
                assert not any("window://office4ai/word" in u for u in uris)
                assert not any("window://office4ai/ppt" in u for u in uris)
                # skill:// 根 + 子资源出现
                assert DEMO_BASE in uris
                assert f"{DEMO_BASE}/SKILL.md" in uris
                assert f"{DEMO_BASE}/assets/logo.png" in uris

                # 根带 _meta.source=resources
                root = next(r for r in result.resources if str(r.uri) == DEMO_BASE)
                assert root.meta is not None and root.meta.get("source") == "resources"

    async def test_real_server_reads_skill_md(self) -> None:
        async with stdio_client(self._server_params()) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.read_resource(AnyUrl(f"{DEMO_BASE}/SKILL.md"))
                content = result.contents[0]
                assert isinstance(content, TextResourceContents)
                assert "demo-authoring" in content.text
