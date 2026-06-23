"""Unit tests for Word OOXML MCP tools (word_get_ooxml / word_insert_ooxml).

测试策略 (镜像 PPT slide-OOXML 工具单测):
- get: 导出的 OOXML 字符串落盘到 dest_path，返回句柄 {scope, filePath, bytes}，不回 inline
- insert: source_path 被读出 → 作为 wire ``ooxml`` 字符串发送 (非 base64、非 source_path)
- scope 回生效值 (空选区回退 body 时 data.scope == "body")
- 错误路径: 文件不存在 / 空文件 / 非法 insertLocation 时，绝不到达 Add-In
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.word import WordGetOoxmlTool, WordInsertOoxmlTool
from office4ai.environment.workspace.base import OfficeObs

SAMPLE_OOXML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pkg:package '
    'xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">body</pkg:package>'
)


@pytest.fixture
def mock_workspace():
    """创建 mock OfficeWorkspace"""
    workspace = MagicMock()
    workspace.execute = AsyncMock()
    return workspace


class TestMetadata:
    def test_get_metadata(self, mock_workspace) -> None:
        tool = WordGetOoxmlTool(mock_workspace)
        assert tool.name == "word_get_ooxml"
        assert tool.category == "word"
        assert tool.event_name == "get:ooxml"
        assert "document_uri" in tool.input_schema["properties"]

    def test_insert_metadata(self, mock_workspace) -> None:
        tool = WordInsertOoxmlTool(mock_workspace)
        assert tool.name == "word_insert_ooxml"
        assert tool.category == "word"
        assert tool.event_name == "insert:ooxml"
        assert "document_uri" in tool.input_schema["properties"]
        # alias advertised on the wire-facing schema
        assert "insertLocation" in tool.input_schema["properties"]


class TestWordGetOoxml:
    @pytest.mark.asyncio
    async def test_writes_ooxml_to_disk_and_returns_handle(self, mock_workspace, tmp_path) -> None:
        """导出的 OOXML 字符串落盘到 dest_path，返回 {scope, filePath, bytes}，不含 ooxml。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"scope": "body", "ooxml": SAMPLE_OOXML})
        dest = tmp_path / "out" / "frag.xml"
        tool = WordGetOoxmlTool(mock_workspace)

        result = await tool.execute({"document_uri": "file:///t.docx", "scope": "body", "dest_path": str(dest)})

        assert result["success"] is True
        assert "ooxml" not in result["data"]  # ooxml 不回传 inline
        assert result["data"]["scope"] == "body"
        assert result["data"]["filePath"] == str(dest)
        assert result["data"]["bytes"] == len(SAMPLE_OOXML.encode("utf-8"))
        # 已落盘且可再读
        assert dest.read_text(encoding="utf-8") == SAMPLE_OOXML

    @pytest.mark.asyncio
    async def test_emits_get_ooxml_with_default_scope(self, mock_workspace, tmp_path) -> None:
        """缺省 scope=selection，作为 wire 参数发出。"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True, data={"scope": "selection", "ooxml": SAMPLE_OOXML}
        )
        tool = WordGetOoxmlTool(mock_workspace)

        await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(tmp_path / "f.xml")})

        action = mock_workspace.execute.call_args.args[0]
        assert action.category == "word"
        assert action.action_name == "get:ooxml"
        assert action.params["scope"] == "selection"

    @pytest.mark.asyncio
    async def test_effective_scope_reported_on_fallback(self, mock_workspace, tmp_path) -> None:
        """请求 selection 但 Add-In 回退 body 时，data.scope 回生效值 'body'。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"scope": "body", "ooxml": SAMPLE_OOXML})
        tool = WordGetOoxmlTool(mock_workspace)

        result = await tool.execute(
            {"document_uri": "file:///t.docx", "scope": "selection", "dest_path": str(tmp_path / "f.xml")}
        )

        assert result["data"]["scope"] == "body"

    @pytest.mark.asyncio
    async def test_failure_passthrough(self, mock_workspace, tmp_path) -> None:
        """Add-In 失败 (如 3016) 原样透传，不落盘。"""
        mock_workspace.execute.return_value = OfficeObs(success=False, data={}, error="3016: not supported")
        dest = tmp_path / "f.xml"
        tool = WordGetOoxmlTool(mock_workspace)

        result = await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(dest)})

        assert result["success"] is False
        assert "3016" in result["error"]
        assert not dest.exists()

    @pytest.mark.asyncio
    async def test_missing_ooxml_in_response_fails(self, mock_workspace, tmp_path) -> None:
        """导出成功但响应无 ooxml 字段时失败。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"scope": "body"})
        tool = WordGetOoxmlTool(mock_workspace)

        result = await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(tmp_path / "f.xml")})

        assert result["success"] is False


class TestWordInsertOoxml:
    @pytest.mark.asyncio
    async def test_reads_file_and_sends_ooxml_string(self, mock_workspace, tmp_path) -> None:
        """source_path 被读出，作为 wire ``ooxml`` 字符串发送 (非 base64、source_path 不上 wire)。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})
        src = tmp_path / "frag.xml"
        src.write_text(SAMPLE_OOXML, encoding="utf-8")
        tool = WordInsertOoxmlTool(mock_workspace)

        result = await tool.execute(
            {
                "document_uri": "file:///t.docx",
                "source_path": str(src),
                "insertLocation": "Replace",
                "scope": "body",
            }
        )

        assert result["success"] is True
        params = mock_workspace.execute.call_args.args[0].params
        assert "source_path" not in params  # source_path 不上 wire
        assert params["ooxml"] == SAMPLE_OOXML  # 原样字符串，非 base64
        assert params["insertLocation"] == "Replace"
        assert params["scope"] == "body"
        assert mock_workspace.execute.call_args.args[0].action_name == "insert:ooxml"

    @pytest.mark.asyncio
    async def test_scope_defaults_to_selection(self, mock_workspace, tmp_path) -> None:
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})
        src = tmp_path / "frag.xml"
        src.write_text(SAMPLE_OOXML, encoding="utf-8")
        tool = WordInsertOoxmlTool(mock_workspace)

        await tool.execute({"document_uri": "file:///t.docx", "source_path": str(src), "insertLocation": "Start"})

        params = mock_workspace.execute.call_args.args[0].params
        assert params["scope"] == "selection"

    @pytest.mark.asyncio
    async def test_missing_file_never_reaches_addin(self, mock_workspace) -> None:
        tool = WordInsertOoxmlTool(mock_workspace)

        result = await tool.execute(
            {"document_uri": "file:///t.docx", "source_path": "/no/such/frag.xml", "insertLocation": "End"}
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_empty_file_rejected(self, mock_workspace, tmp_path) -> None:
        src = tmp_path / "empty.xml"
        src.write_text("   \n", encoding="utf-8")
        tool = WordInsertOoxmlTool(mock_workspace)

        result = await tool.execute(
            {"document_uri": "file:///t.docx", "source_path": str(src), "insertLocation": "End"}
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_insert_location_rejected(self, mock_workspace, tmp_path) -> None:
        """非法 insertLocation 在校验阶段失败，绝不到达 Add-In。"""
        src = tmp_path / "frag.xml"
        src.write_text(SAMPLE_OOXML, encoding="utf-8")
        tool = WordInsertOoxmlTool(mock_workspace)

        result = await tool.execute(
            {"document_uri": "file:///t.docx", "source_path": str(src), "insertLocation": "Before"}
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()
