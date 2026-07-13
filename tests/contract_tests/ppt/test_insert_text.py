"""
Contract Tests for ppt:insert:text

测试 ppt:insert:text 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_text_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功插入文本。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert request["text"] == "Hello PPT"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_text_response(element_id="shape-015", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:text", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:text",
            params={"document_uri": client.document_uri, "text": "Hello PPT"},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["elementId"] == "shape-015"
        assert result.data["slideIndex"] == 0

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:insert:text"
        assert event_data["text"] == "Hello PPT"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_text_with_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试带选项插入文本（位置/字体）。"""
    expected_options = {
        "slideIndex": 2,
        "left": 100.0,
        "top": 200.0,
        "width": 300.0,
        "height": 50.0,
        "font": {"size": 24, "name": "Arial", "color": "#FF0000", "strikethrough": True},
    }

    def response_factory(request: dict) -> dict:
        assert "options" in request
        opts = request["options"]
        assert opts["slideIndex"] == 2
        # OASP 0.4.0: 插入即带字体，字体收敛到 font 子对象
        assert opts["font"]["size"] == 24
        assert opts["font"]["color"] == "#FF0000"
        assert opts["font"]["strikethrough"] is True

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_text_response(element_id="shape-016", slide_index=2),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:text", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:text",
            params={
                "document_uri": client.document_uri,
                "text": "Styled text",
                "options": expected_options,
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["slideIndex"] == 2
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_text_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Insert failed"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:text", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:text",
            params={"document_uri": client.document_uri, "text": "Fail"},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
