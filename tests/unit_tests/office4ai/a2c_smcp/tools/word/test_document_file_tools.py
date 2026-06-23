"""Unit tests for Word whole-document .docx MCP tools (word_get_document_file / word_insert_document_file).

测试策略 (镜像 PPT slide-OOXML 工具单测 + 本仓 ooxml 工具单测):
- get: 导出的 base64 解码落盘 dest_path (.docx 二进制)，返回 {filePath, bytes}，不回 inline base64
- insert: source_path 字节被读出 → base64 编码后作为 wire base64 发送 (source_path 不上 wire)
- 错误路径: 文件不存在 / 空文件 / 非法 insertLocation 时，绝不到达 Add-In
"""

from __future__ import annotations

import base64
from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.word import WordGetDocumentFileTool, WordInsertDocumentFileTool
from office4ai.environment.workspace.base import OfficeObs

# 一个最小的「.docx 包字节」替身 (zip 魔数 PK\x03\x04 开头)，仅验证字节 round-trip，不要求真实可解
SAMPLE_DOCX_BYTES = b"PK\x03\x04fake-docx-bytes\x00\x01\x02"
SAMPLE_DOCX_B64 = base64.b64encode(SAMPLE_DOCX_BYTES).decode("ascii")


@pytest.fixture
def mock_workspace():
    """创建 mock OfficeWorkspace"""
    workspace = MagicMock()
    workspace.execute = AsyncMock()
    return workspace


class TestMetadata:
    def test_get_metadata(self, mock_workspace) -> None:
        tool = WordGetDocumentFileTool(mock_workspace)
        assert tool.name == "word_get_document_file"
        assert tool.category == "word"
        assert tool.event_name == "get:documentFile"
        assert "document_uri" in tool.input_schema["properties"]

    def test_insert_metadata(self, mock_workspace) -> None:
        tool = WordInsertDocumentFileTool(mock_workspace)
        assert tool.name == "word_insert_document_file"
        assert tool.category == "word"
        assert tool.event_name == "insert:documentFile"
        assert "document_uri" in tool.input_schema["properties"]
        assert "insertLocation" in tool.input_schema["properties"]


class TestWordGetDocumentFile:
    @pytest.mark.asyncio
    async def test_decodes_base64_to_disk_and_returns_handle(self, mock_workspace, tmp_path) -> None:
        """导出的 base64 解码落盘到 dest_path，返回 {filePath, bytes}，不含 base64。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"base64": SAMPLE_DOCX_B64})
        dest = tmp_path / "out" / "doc.docx"
        tool = WordGetDocumentFileTool(mock_workspace)

        result = await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(dest)})

        assert result["success"] is True
        assert "base64" not in result["data"]  # base64 不回传 inline
        assert result["data"]["filePath"] == str(dest)
        assert result["data"]["bytes"] == len(SAMPLE_DOCX_BYTES)
        # 已落盘且字节一致 (可被解 zip)
        assert dest.read_bytes() == SAMPLE_DOCX_BYTES

    @pytest.mark.asyncio
    async def test_emits_get_document_file(self, mock_workspace, tmp_path) -> None:
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"base64": SAMPLE_DOCX_B64})
        tool = WordGetDocumentFileTool(mock_workspace)
        await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(tmp_path / "d.docx")})
        action = mock_workspace.execute.call_args.args[0]
        assert action.category == "word"
        assert action.action_name == "get:documentFile"
        # whole-document export: no scope on the wire
        assert "scope" not in action.params

    @pytest.mark.asyncio
    async def test_failure_passthrough_no_write(self, mock_workspace, tmp_path) -> None:
        mock_workspace.execute.return_value = OfficeObs(success=False, data={}, error="3004: nope")
        dest = tmp_path / "d.docx"
        tool = WordGetDocumentFileTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(dest)})
        assert result["success"] is False
        assert not dest.exists()

    @pytest.mark.asyncio
    async def test_missing_base64_in_response_fails(self, mock_workspace, tmp_path) -> None:
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})
        tool = WordGetDocumentFileTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///t.docx", "dest_path": str(tmp_path / "d.docx")})
        assert result["success"] is False


class TestWordInsertDocumentFile:
    @pytest.mark.asyncio
    async def test_reads_file_and_sends_base64(self, mock_workspace, tmp_path) -> None:
        """source_path 字节被读出、base64 编码后作为 wire base64 发送 (source_path 不上 wire)。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})
        src = tmp_path / "doc.docx"
        src.write_bytes(SAMPLE_DOCX_BYTES)
        tool = WordInsertDocumentFileTool(mock_workspace)

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
        # base64 可解回原始字节
        assert base64.b64decode(params["base64"]) == SAMPLE_DOCX_BYTES
        assert params["insertLocation"] == "Replace"
        assert params["scope"] == "body"
        assert mock_workspace.execute.call_args.args[0].action_name == "insert:documentFile"

    @pytest.mark.asyncio
    async def test_scope_defaults_to_body(self, mock_workspace, tmp_path) -> None:
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})
        src = tmp_path / "doc.docx"
        src.write_bytes(SAMPLE_DOCX_BYTES)
        tool = WordInsertDocumentFileTool(mock_workspace)
        await tool.execute({"document_uri": "file:///t.docx", "source_path": str(src), "insertLocation": "End"})
        params = mock_workspace.execute.call_args.args[0].params
        assert params["scope"] == "body"

    @pytest.mark.asyncio
    async def test_missing_file_never_reaches_addin(self, mock_workspace) -> None:
        tool = WordInsertDocumentFileTool(mock_workspace)
        result = await tool.execute(
            {"document_uri": "file:///t.docx", "source_path": "/no/such/doc.docx", "insertLocation": "End"}
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_empty_file_rejected(self, mock_workspace, tmp_path) -> None:
        src = tmp_path / "empty.docx"
        src.write_bytes(b"")
        tool = WordInsertDocumentFileTool(mock_workspace)
        result = await tool.execute(
            {"document_uri": "file:///t.docx", "source_path": str(src), "insertLocation": "End"}
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_insert_location_rejected(self, mock_workspace, tmp_path) -> None:
        src = tmp_path / "doc.docx"
        src.write_bytes(SAMPLE_DOCX_BYTES)
        tool = WordInsertDocumentFileTool(mock_workspace)
        result = await tool.execute(
            {"document_uri": "file:///t.docx", "source_path": str(src), "insertLocation": "Before"}
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()
