"""
Contract Tests for ppt:goto:slide

测试 ppt:goto:slide 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace, parse_error_details


@pytest.mark.asyncio
@pytest.mark.contract
async def test_goto_slide_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功跳转到幻灯片。"""

    def response_factory(request: dict) -> dict:
        assert request["slideIndex"] == 5

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.goto_slide_response(current_index=5),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:goto:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="goto:slide",
            params={"document_uri": client.document_uri, "slideIndex": 5},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["currentSlideIndex"] == 5

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:goto:slide"
        assert event_data["slideIndex"] == 5
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_goto_slide_index_out_of_range_3008(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """slideIndex 越界 → 3008 POSITION_INVALID + details（oasp#23 / issue #90）。

    与 ``ppt:delete:slide`` 同判法：线缆层合法、仅相对当前文档状态无效，恢复靠「重读后重试」。
    oasp#23 仅取 /ppt 这两个事件作样板验证判法，全量清扫见 oasp#24——**断言不得扩大化**到
    其他序号越界事件，其余沿用各自事件表内现码。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3008",
                "message": "Slide index out of range",
                "details": {"index": 999, "total": 10, "kind": "slide"},
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:goto:slide", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="goto:slide",
            params={"document_uri": client.document_uri, "slideIndex": 999},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3008" in (result.error or "")
        details = parse_error_details(result.error or "")
        assert details == {"index": 999, "total": 10, "kind": "slide"}
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_goto_slide_verify_index(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试跳转到第一张幻灯片。"""

    def response_factory(request: dict) -> dict:
        assert request["slideIndex"] == 0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.goto_slide_response(current_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:goto:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="goto:slide",
            params={"document_uri": client.document_uri, "slideIndex": 0},
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["currentSlideIndex"] == 0
    finally:
        await client.disconnect()
