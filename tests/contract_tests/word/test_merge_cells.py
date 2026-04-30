"""
Contract Tests for word:merge:cells (OASP /word Draft, v0.2.0)

测试 word:merge:cells 事件的请求-响应流程。

OASP Spec: https://doc.turingfocus.cn/oasp/0.2.0/specification/events-word/
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_merge_cells_with_explicit_table_id(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """显式 tableId 路径：合并首行 5 列为整行表头单元格。"""

    def response_factory(request: dict) -> dict:
        # Wire payload uses camelCase per OASP protocol
        assert "requestId" in request
        assert "documentUri" in request
        assert request["tableId"] == "table-0"
        assert request["startRowIndex"] == 0
        assert request["startColumnIndex"] == 0
        assert request["endRowIndex"] == 0
        assert request["endColumnIndex"] == 4

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.merge_cells_response(table_id="table-0", row_count=1, column_count=5),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:merge:cells", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="merge:cells",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 4,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["tableId"] == "table-0"
        assert result.data["requestedRange"] == {"rowCount": 1, "columnCount": 5}

        assert len(client.received_events) == 1
        event_name, payload = client.received_events[0]
        assert event_name == "word:merge:cells"
        # tableId on the wire — confirms snake_case → camelCase conversion at DTO layer
        assert payload["tableId"] == "table-0"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_merge_cells_omitted_table_id(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """缺省 tableId 路径：服务端不应在 wire payload 中包含 tableId 键。"""

    def response_factory(request: dict) -> dict:
        # Optional tableId must be omitted when not provided (exclude_none on DTO)
        assert "tableId" not in request, "tableId should be omitted when not provided"
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.merge_cells_response(table_id="table-3", row_count=1, column_count=3),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:merge:cells", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="merge:cells",
            params={
                "document_uri": client.document_uri,
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 2,
            },
        )
        result = await workspace.execute(action)
        assert result.success is True
        # Add-In resolved tableId from cursor — server propagates the resolved value
        assert result.data["tableId"] == "table-3"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_merge_cells_no_table_at_cursor(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """缺省 tableId + 光标不在表格内 → 3013 NO_TABLE_AT_CURSOR."""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3013",
                "message": "Cursor is not inside a table; specify tableId explicitly.",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:merge:cells", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="merge:cells",
            params={
                "document_uri": client.document_uri,
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 2,
            },
        )
        result = await workspace.execute(action)
        assert result.success is False
        assert "3013" in (result.error or "")
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_merge_cells_already_merged(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """目标区域已存在合并冲突 → 3014 ALREADY_MERGED."""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3014",
                "message": "Target range conflicts with an existing merged region.",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:merge:cells", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="merge:cells",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 1,
                "end_column_index": 2,
            },
        )
        result = await workspace.execute(action)
        assert result.success is False
        assert "3014" in (result.error or "")
    finally:
        await client.disconnect()
