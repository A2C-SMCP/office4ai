"""
Contract Tests for ppt:delete:slide

测试 ppt:delete:slide 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace, parse_error_details


@pytest.mark.asyncio
@pytest.mark.contract
async def test_delete_slide_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功删除幻灯片。"""

    def response_factory(request: dict) -> dict:
        assert request["slideIndex"] == 3

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.delete_slide_response(deleted_index=3, new_count=9),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:delete:slide", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="delete:slide",
            params={"document_uri": client.document_uri, "slideIndex": 3},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["deletedIndex"] == 3
        assert result.data["newSlideCount"] == 9

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:delete:slide"
        assert event_data["slideIndex"] == 3
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_delete_slide_index_out_of_range_3008(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """slideIndex 越界 → 3008 POSITION_INVALID + details（oasp#23 / issue #90）。

    该值在线缆上完全合法（非负整数、类型正确），仅相对**当前文档状态**无效，故不是 4002/4004——
    正确恢复是「重读文档状态后原值重试」而非「改参数」。本用例钉住 Server 对权威码与 details
    的无损透传；码由 Add-In 发出（配套 office-editor4ai#92）。
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

    client.register_response("ppt:delete:slide", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="delete:slide",
            params={"document_uri": client.document_uri, "slideIndex": 999},
        )
        result = await workspace.execute(action)

        assert result.success is False
        # 权威码必须原样到达调用方——只断言 success is False 等于对错误码不设防
        assert "3008" in (result.error or "")
        # details 承载「请求序号 + 实时边界」，调用方据此决定重读后重试；须无损透传
        details = parse_error_details(result.error or "")
        assert details == {"index": 999, "total": 10, "kind": "slide"}
    finally:
        await client.disconnect()
