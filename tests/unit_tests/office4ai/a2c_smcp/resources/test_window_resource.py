"""WindowResource 根索引 单元测试 | WindowResource root index unit tests"""

from __future__ import annotations

from unittest.mock import patch

import pytest

from office4ai.a2c_smcp.resources.window import WindowResource
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from tests.unit_tests.office4ai.a2c_smcp.resources.conftest import make_client


@pytest.fixture
def resource(workspace: OfficeWorkspace) -> WindowResource:
    return WindowResource(workspace, priority=0, fullscreen=False)


class TestWindowResourceMetadata:
    def test_uri_format(self, resource: WindowResource) -> None:
        assert resource.uri == "window://office4ai?priority=0&fullscreen=false"

    def test_base_uri(self, resource: WindowResource) -> None:
        assert resource.base_uri == "window://office4ai"

    def test_name(self, resource: WindowResource) -> None:
        assert resource.name == "Office 工作区"

    def test_mime_type(self, resource: WindowResource) -> None:
        assert resource.mime_type == "text/plain"

    def test_priority_validation(self) -> None:
        ws = OfficeWorkspace.__new__(OfficeWorkspace)
        ws._last_activity = None
        ws._content_cache = {}
        ws._structure_cache = {}
        with pytest.raises(ValueError, match="priority must be int"):
            WindowResource(ws, priority=-1)
        with pytest.raises(ValueError, match="priority must be int"):
            WindowResource(ws, priority=101)


class TestWindowResourceRead:
    @pytest.mark.asyncio
    async def test_read_empty(self, resource: WindowResource) -> None:
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = []
            content = await resource.read()

        assert "Office 工作区" in content
        assert "暂无文档连接" in content

    @pytest.mark.asyncio
    async def test_read_word_only(self, resource: WindowResource) -> None:
        # W4b-1：根索引逐条列出每文件的 per-file window 子资源
        clients = [
            make_client("file:///tmp/a.docx", namespace="/word", socket_id="s1"),
            make_client("file:///tmp/b.docx", namespace="/word", socket_id="s2", client_id="c2"),
        ]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "WORD · a.docx" in content
        assert "WORD · b.docx" in content
        assert "window://office4ai/word/" in content
        assert "PPT" not in content

    @pytest.mark.asyncio
    async def test_read_ppt_only(self, resource: WindowResource) -> None:
        clients = [
            make_client("file:///tmp/slides.pptx", namespace="/ppt"),
        ]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "PPT · slides.pptx" in content
        assert "window://office4ai/ppt/" in content
        assert "WORD" not in content

    @pytest.mark.asyncio
    async def test_read_mixed(self, resource: WindowResource) -> None:
        clients = [
            make_client("file:///tmp/a.docx", namespace="/word", socket_id="s1"),
            make_client("file:///tmp/slides.pptx", namespace="/ppt", socket_id="s2", client_id="c2"),
        ]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "WORD · a.docx" in content
        assert "PPT · slides.pptx" in content

    @pytest.mark.asyncio
    async def test_excel_skipped_until_w4b3(self, resource: WindowResource) -> None:
        """excel 连接暂无 per-file 窗口（W4b-3 延后）→ 根索引不列出。"""
        clients = [make_client("file:///tmp/book.xlsx", namespace="/excel")]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        assert "book.xlsx" not in content
        assert "暂无文档连接" in content

    @pytest.mark.asyncio
    async def test_dedup_by_uri(self, resource: WindowResource) -> None:
        """Same document with multiple socket connections should be listed once."""
        clients = [
            make_client("file:///tmp/a.docx", namespace="/word", socket_id="s1"),
            make_client("file:///tmp/a.docx", namespace="/word", socket_id="s2", client_id="c2"),
        ]
        with patch("office4ai.a2c_smcp.resources.window.connection_manager") as mock_cm:
            mock_cm.get_all_clients.return_value = clients
            content = await resource.read()

        # one window line for the doc despite two socket connections
        assert content.count("WORD · a.docx") == 1


class TestWindowResourceUpdateFromUri:
    def test_update_from_uri_priority(self, resource: WindowResource) -> None:
        resource.update_from_uri("window://office4ai?priority=50&fullscreen=false")
        assert "priority=50" in resource.uri

    def test_update_from_uri_fullscreen(self, resource: WindowResource) -> None:
        resource.update_from_uri("window://office4ai?priority=0&fullscreen=true")
        assert "fullscreen=true" in resource.uri

    def test_update_from_uri_invalid_priority(self, resource: WindowResource) -> None:
        resource.update_from_uri("window://office4ai?priority=200&fullscreen=false")
        assert "priority=0" in resource.uri
