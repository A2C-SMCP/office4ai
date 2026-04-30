"""
Contract Tests for word:update:tableRowColumn (OASP /word Draft, v0.2.0)

测试 word:update:tableRowColumn 事件的请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_rows(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """按行批量写入 4 行数据。"""

    def response_factory(request: dict) -> dict:
        assert request["tableId"] == "table-0"
        assert len(request["rows"]) == 4
        # First row content sanity
        assert request["rows"][0]["rowIndex"] == 1
        assert request["rows"][0]["values"] == ["甲方", "ACME Corp"]
        # columns key omitted via exclude_none
        assert "columns" not in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.update_table_row_column_response(
                table_id="table-0", cells_updated=8, row_count=5, column_count=4
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableRowColumn", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "rows": [
                    {"rowIndex": 1, "values": ["甲方", "ACME Corp"]},
                    {"rowIndex": 2, "values": ["地址", "上海市"]},
                    {"rowIndex": 3, "values": ["联系人", "张三"]},
                    {"rowIndex": 4, "values": ["日期", "2026-04-30"]},
                ],
            },
        )
        result = await workspace.execute(action)
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["cellsUpdated"] == 8
        assert result.data["rowCount"] == 5
        assert result.data["columnCount"] == 4
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_columns(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """按列批量写入。"""

    def response_factory(request: dict) -> dict:
        assert "rows" not in request
        assert len(request["columns"]) == 1
        assert request["columns"][0]["columnIndex"] == 0
        assert request["columns"][0]["values"] == ["甲方", "地址", "联系人"]

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.update_table_row_column_response(
                table_id="table-0", cells_updated=3, row_count=5, column_count=4
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableRowColumn", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "columns": [
                    {"columnIndex": 0, "values": ["甲方", "地址", "联系人"]},
                ],
            },
        )
        result = await workspace.execute(action)
        assert result.success is True
        assert result.data["cellsUpdated"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_row_column_no_table_at_cursor(
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

    client.register_response("word:update:tableRowColumn", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableRowColumn",
            params={
                "document_uri": client.document_uri,
                "rows": [{"rowIndex": 0, "values": ["x", "y"]}],
            },
        )
        result = await workspace.execute(action)
        assert result.success is False
        assert "3013" in (result.error or "")
    finally:
        await client.disconnect()
