"""
Contract Tests for word:update:tableCell (OASP /word Draft, v0.2.0)

测试 word:update:tableCell 事件的请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_cell_with_format(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """正常路径：单元格文本 + 完整 CellFormat（蓝底居中白字加粗表头）。"""

    def response_factory(request: dict) -> dict:
        assert request["tableId"] == "table-0"
        assert len(request["cells"]) == 1
        cell = request["cells"][0]
        assert cell["rowIndex"] == 0
        assert cell["columnIndex"] == 0
        assert cell["text"] == "甲方信息"
        # CellFormat fields must hit the wire in camelCase
        fmt = cell["format"]
        assert fmt["horizontalAlignment"] == "Centered"
        assert fmt["verticalAlignment"] == "Center"
        assert fmt["backgroundColor"] == "#1F4E79"
        assert fmt["fontColor"] == "#FFFFFF"
        assert fmt["bold"] is True

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.update_table_cell_response(
                table_id="table-0", cells_updated=1, row_count=5, column_count=4
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableCell", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableCell",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "cells": [
                    {
                        "rowIndex": 0,
                        "columnIndex": 0,
                        "text": "甲方信息",
                        "format": {
                            "horizontalAlignment": "Centered",
                            "verticalAlignment": "Center",
                            "backgroundColor": "#1F4E79",
                            "fontColor": "#FFFFFF",
                            "bold": True,
                        },
                    }
                ],
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["tableId"] == "table-0"
        assert result.data["cellsUpdated"] == 1
        assert result.data["rowCount"] == 5
        assert result.data["columnCount"] == 4

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "word:update:tableCell"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_cell_batch(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """批量更新多个单元格（标签/填写区两列结构）。"""

    def response_factory(request: dict) -> dict:
        assert len(request["cells"]) == 4
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.update_table_cell_response(
                table_id="table-0", cells_updated=4, row_count=5, column_count=4
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableCell", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableCell",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "cells": [
                    {
                        "rowIndex": 1,
                        "columnIndex": 0,
                        "text": "甲方",
                        "format": {"backgroundColor": "#EEEEEE", "bold": True},
                    },
                    {"rowIndex": 1, "columnIndex": 1, "text": "ACME Corp"},
                    {
                        "rowIndex": 2,
                        "columnIndex": 0,
                        "text": "地址",
                        "format": {"backgroundColor": "#EEEEEE", "bold": True},
                    },
                    {"rowIndex": 2, "columnIndex": 1, "text": "上海市"},
                ],
            },
        )
        result = await workspace.execute(action)
        assert result.success is True
        assert result.data["cellsUpdated"] == 4
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_cell_table_not_found(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """tableId 不存在 → 3010 ELEMENT_NOT_FOUND."""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3010", "message": "tableId 'table-99' not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableCell", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableCell",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-99",
                "cells": [{"rowIndex": 0, "columnIndex": 0, "text": "x"}],
            },
        )
        result = await workspace.execute(action)
        assert result.success is False
        assert "3010" in (result.error or "")
    finally:
        await client.disconnect()
