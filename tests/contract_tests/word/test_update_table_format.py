"""
Contract Tests for word:update:tableFormat (OASP /word Draft, v0.2.0)

测试 word:update:tableFormat 事件的请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_full(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """整表样式 + 内边距 + 边框 + 列宽 + 对齐。"""

    def response_factory(request: dict) -> dict:
        assert request["tableId"] == "table-0"
        assert request["alignment"] == "Centered"
        assert request["columnWidths"] == [120, 80, 80, 80]
        # styleOptions / borderOptions 都用 camelCase
        assert request["styleOptions"]["styleType"] == "Grid Table 4 - Accent 1"
        assert request["styleOptions"]["bandedRows"] is True
        assert request["styleOptions"]["cellPadding"] == {
            "top": 4,
            "bottom": 4,
            "left": 6,
            "right": 6,
        }
        assert request["borderOptions"]["location"] == "inside"
        assert request["borderOptions"]["style"] == "Single"
        assert request["borderOptions"]["width"] == 0.5

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.update_table_format_response(table_id="table-0", row_count=5, column_count=4),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableFormat", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "style_options": {
                    "styleType": "Grid Table 4 - Accent 1",
                    "bandedRows": True,
                    "cellPadding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
                },
                "border_options": {
                    "location": "inside",
                    "style": "Single",
                    "width": 0.5,
                },
                "column_widths": [120, 80, 80, 80],
                "alignment": "Centered",
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["tableId"] == "table-0"
        assert result.data["rowCount"] == 5
        assert result.data["columnCount"] == 4

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "word:update:tableFormat"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_borders_only(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """仅边框：清除外框 (style='None') 验证枚举值。"""

    def response_factory(request: dict) -> dict:
        # 仅 borderOptions 出现，其它可选键被 exclude_none 过滤
        assert "styleOptions" not in request
        assert "columnWidths" not in request
        assert "alignment" not in request
        assert request["borderOptions"]["location"] == "outside"
        assert request["borderOptions"]["style"] == "None"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.update_table_format_response(row_count=5, column_count=4),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableFormat", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "border_options": {"location": "outside", "style": "None"},
            },
        )
        result = await workspace.execute(action)
        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_style_not_found(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """styleType 不存在 → 3011 STYLE_NOT_FOUND."""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3011",
                "message": "The style does not exist.",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableFormat", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "style_options": {"styleType": "NoSuchStyle"},
            },
        )
        result = await workspace.execute(action)
        assert result.success is False
        assert "3011" in (result.error or "")
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_column_widths_too_long(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """columnWidths 长度超出列数 → 4002 INVALID_PARAM (Add-In 校验)."""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "4002",
                "message": "columnWidths length exceeds table column count",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:update:tableFormat", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "table_id": "table-0",
                "column_widths": [60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60],
            },
        )
        result = await workspace.execute(action)
        assert result.success is False
        assert "4002" in (result.error or "")
    finally:
        await client.disconnect()
