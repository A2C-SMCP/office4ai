"""
Contract Tests for ppt:insert:shape

测试 ppt:insert:shape 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功插入形状。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert request["shapeType"] == "Rectangle"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_shape_response(shape_id="shape-020", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={"document_uri": client.document_uri, "shapeType": "Rectangle"},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["elementId"] == "shape-020"

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:insert:shape"
        assert event_data["shapeType"] == "Rectangle"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_different_types(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试不同形状类型。"""

    def response_factory(request: dict) -> dict:
        assert request["shapeType"] == "Circle"
        assert "options" in request
        assert request["options"]["fillColor"] == "#0000FF"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_shape_response(shape_id="shape-021", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={
                "document_uri": client.document_uri,
                "shapeType": "Circle",
                "options": {
                    "left": 200.0,
                    "top": 150.0,
                    "width": 100.0,
                    "height": 100.0,
                    "fillColor": "#0000FF",
                },
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["elementId"] == "shape-021"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Shape insert failed"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={"document_uri": client.document_uri, "shapeType": "Triangle"},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_with_font(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """OASP 0.4.0: 插入即带字体 — font 子对象作为 camelCase wire payload 透传到 AddIn。

    text-capable 形状（此处 RoundedRectangle）随 text + font 一次性插入并带格式。
    """

    def response_factory(request: dict) -> dict:
        assert request["shapeType"] == "RoundedRectangle"
        opts = request["options"]
        # text → font 施加顺序：text 与 font 并存透传
        assert opts["text"] == "点击这里"
        # font 收敛为子对象，wire 为 camelCase
        assert opts["font"]["size"] == 18
        assert opts["font"]["name"] == "微软雅黑"
        assert opts["font"]["color"] == "#FFFFFF"
        assert opts["font"]["bold"] is True
        # snake_case 入参 → camelCase wire（load-bearing：double_strikethrough != doubleStrikethrough）
        assert opts["font"]["doubleStrikethrough"] is True
        assert "double_strikethrough" not in opts["font"]

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.insert_shape_response(shape_id="shape-022", slide_index=0),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={
                "document_uri": client.document_uri,
                "shapeType": "RoundedRectangle",
                "options": {
                    "text": "点击这里",
                    "font": {
                        "size": 18,
                        "name": "微软雅黑",
                        "color": "#FFFFFF",
                        "bold": True,
                        "double_strikethrough": True,
                    },
                },
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["elementId"] == "shape-022"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_shape_line_font_rejected_by_addin(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """text-capable 门控（透传口径）：Server 不本地校验，font/text 照发到无文本框的 Line；

    text-capable 规则由 AddIn 权威裁决，前置静态拒绝返回 4002 —— Server 透传并兜住该返回码。
    """

    def line_reject_response(request: dict) -> dict:
        # Server 透传：text / font 未被本地拦截，照原样送达 AddIn（薄传输，text/font 一视同仁）
        assert request["shapeType"] == "Line"
        assert "font" in request["options"]
        assert request["options"]["text"] == "on a line"
        # AddIn 静态拒绝（语义误用 4002，与能力不足 3016 分离）
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "4002", "message": "font/text not applicable to a shape without a text box (Line)"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:insert:shape", line_reject_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="insert:shape",
            params={
                "document_uri": client.document_uri,
                "shapeType": "Line",
                "options": {"text": "on a line", "font": {"size": 18, "bold": True}},
            },
        )
        result = await workspace.execute(action)

        # 透传的 4002 被 Server 兜住并展开为 "4002: ..." 字符串（office_workspace flatten）
        assert result.success is False
        assert result.error is not None
        assert "4002" in result.error
    finally:
        await client.disconnect()
