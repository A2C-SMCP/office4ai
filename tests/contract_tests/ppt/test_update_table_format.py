"""
Contract Tests for ppt:update:tableFormat

测试 ppt:update:tableFormat 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_cell_formats(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试单元格级别格式更新。"""
    cell_formats = [
        {
            "rowIndex": 0,
            "columnIndex": 0,
            "backgroundColor": "#FF0000",
            "font": {"bold": True, "size": 14, "underline": "Single"},
            "horizontalAlignment": "Center",
            "verticalAlignment": "Middle",
        },
    ]

    def response_factory(request: dict) -> dict:
        assert request["elementId"] == "table-001"
        assert len(request["cellFormats"]) == 1
        fmt = request["cellFormats"][0]
        assert fmt["backgroundColor"] == "#FF0000"
        # OASP 0.4.0: 字体收敛到 font 子对象
        assert fmt["font"]["bold"] is True
        assert fmt["font"]["size"] == 14
        assert fmt["font"]["underline"] == "Single"
        # 对齐改用 office.js PowerPoint 枚举（Center/Middle）
        assert fmt["horizontalAlignment"] == "Center"
        assert fmt["verticalAlignment"] == "Middle"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_format_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableFormat", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "cellFormats": cell_formats,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 1

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:update:tableFormat"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_row_formats(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试行级别格式更新。"""
    row_formats = [
        {"rowIndex": 0, "height": 50.0, "backgroundColor": "#EEEEEE", "font": {"size": 16, "bold": True}},
    ]

    def response_factory(request: dict) -> dict:
        assert len(request["rowFormats"]) == 1
        assert request["rowFormats"][0]["height"] == 50.0
        # OASP 0.4.0: 行级字体升级为完整 PptFont（原仅 fontSize）
        assert request["rowFormats"][0]["font"]["size"] == 16
        assert request["rowFormats"][0]["font"]["bold"] is True

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_format_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableFormat", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "rowFormats": row_formats,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_column_formats(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试列级别格式更新。"""
    column_formats = [
        {"columnIndex": 0, "width": 150.0, "backgroundColor": "#DDDDDD"},
        {"columnIndex": 1, "width": 200.0, "font": {"size": 12}},
    ]

    def response_factory(request: dict) -> dict:
        assert len(request["columnFormats"]) == 2
        assert request["columnFormats"][0]["width"] == 150.0
        # OASP 0.4.0: 列级字体升级为完整 PptFont（原仅 fontSize）
        assert request["columnFormats"][1]["font"]["size"] == 12

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_table_format_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableFormat", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "elementId": "table-001",
                "columnFormats": column_formats,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_table_format_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Table not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:tableFormat", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:tableFormat",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "cellFormats": [{"rowIndex": 0, "columnIndex": 0, "font": {"bold": True}}],
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
