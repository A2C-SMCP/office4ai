"""
Contract Tests for ppt:get:slideScreenshot

测试 ppt:get:slideScreenshot 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace, parse_error_details


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功获取幻灯片截图（PNG 格式）。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        assert request["slideIndex"] == 0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_screenshot_response(format="png"),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={"document_uri": client.document_uri, "slideIndex": 0},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["format"] == "png"
        assert len(result.data["base64"]) > 0

        assert len(client.received_events) == 1
        event_name, _ = client.received_events[0]
        assert event_name == "ppt:get:slideScreenshot"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_jpeg_format(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试 JPEG 格式截图。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        assert request["options"]["format"] == "jpeg"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.slide_screenshot_response(format="jpeg"),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={
                "document_uri": client.document_uri,
                "slideIndex": 0,
                "options": {"format": "jpeg", "quality": 80},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
        assert result.data["format"] == "jpeg"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Slide not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={"document_uri": client.document_uri, "slideIndex": 999},
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_format_not_supported_3007(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """format 枚举合法但目标渲染不出该格式 → 3007 FORMAT_NOT_SUPPORTED（oasp#23 / issue #90）。

    输出侧语义：请求的 ``options.format`` 通过了线缆层校验（是枚举内取值），只是宿主产不出
    这一种——恢复动作是**改请求另一格式重发**。原落 ``3000`` 兜底，调用方无从区分。
    与下方 3016 的分界见 ``test_..._api_not_supported_3016``。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3007",
                "message": "Screenshot format 'svg' cannot be rendered on this host",
                "details": {"format": "svg", "accepted": ["png", "jpeg"]},
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={"document_uri": client.document_uri, "slideIndex": 1},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3007" in (result.error or "")
        details = parse_error_details(result.error or "")
        assert details is not None
        assert details["format"] == "svg"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_slide_screenshot_api_not_supported_3016(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """整个截图能力在当前宿主不可用 → 3016 API_NOT_SUPPORTED（oasp#23 / issue #90）。

    与 3007 的分界是**失败粒度**而非「谁的锅」——3016 意味着**换格式无用**，须换路径或平台；
    误报成 3007 会让调用方在一个根本不存在的能力上徒劳试遍所有格式。截图/导出这类高度依赖
    宿主能力的输出侧事件，3016 是其降级路由，故规范按本仓体例给它**独立一行**而非藏在
    3007 行的括号里。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3016",
                "message": "Slide screenshot is not available on this platform",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:get:slideScreenshot", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="get:slideScreenshot",
            params={"document_uri": client.document_uri, "slideIndex": 1},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3016" in (result.error or "")
    finally:
        await client.disconnect()
