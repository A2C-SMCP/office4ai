"""
Contract Tests for ppt:update:image

测试 ppt:update:image 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace, parse_error_details

TEST_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功替换图片内容。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert request["elementId"] == "img-001"
        assert request["image"]["base64"] == TEST_BASE64

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_image_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "img-001",
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 1

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:update:image"
        assert event_data["elementId"] == "img-001"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_keep_dimensions(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试 keepDimensions 选项传递。"""

    def response_factory(request: dict) -> dict:
        assert "options" in request
        assert request["options"]["keepDimensions"] is False
        assert request["options"]["width"] == 500.0
        assert request["options"]["height"] == 300.0

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_image_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "img-001",
                "image": {"base64": TEST_BASE64},
                "options": {"keepDimensions": False, "width": 500.0, "height": 300.0},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_element_not_found_3010(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """elementId 不存在 → 3010 ELEMENT_NOT_FOUND（events-ppt.md §ppt:update:image 错误表）。

    原用例固化 ``3001``（该码在本事件表里专指「文档未找到」），与规范矛盾且断言只查
    ``success is False`` 故未暴露——issue #90 一并订正。

    ``details.kind`` 取 ``element``：该词表是**规范层**（error-handling.md §ELEMENT_NOT_FOUND
    的「双端据此对齐断言」表），``/word``、``/ppt`` 下按 ``elementId`` 定位的内容元素统一用
    ``element``——**没有** ``image`` 这一取值，已列值语义固定、不得另造。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3010",
                "message": "Image element not found",
                "details": {"kind": "element", "id": "nonexistent"},
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3010" in (result.error or "")
        # 断言 details 而非只断言码——否则 mock 里写错的词表值永远不会被发现
        details = parse_error_details(result.error or "")
        assert details is not None
        assert details["kind"] == "element"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_image_format_not_supported_3007(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """base64 可解码但图片格式不受支持 → 3007 FORMAT_NOT_SUPPORTED（oasp#23 / issue #90）。

    与 ``4002 INVALID_PARAM`` 的分界是**失败粒度**：4002 = 载荷在线缆层即不合法（base64 无法
    解码），3007 = 数据完整可解码、只是宿主拒绝这一种格式（**转码后重发即可成功**）。
    再往上 ``3016`` 才是「整个能力不可用、换格式无用」。三码混用会把调用方引向错误的恢复动作。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3007",
                "message": "Image format 'image/tiff' is not supported",
                "details": {"format": "image/tiff", "accepted": ["image/png", "image/jpeg"]},
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:image", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:image",
            params={
                "document_uri": client.document_uri,
                "elementId": "img-1",
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3007" in (result.error or "")
        # details 回带被拒格式与可接受列表，调用方据此一次选对目标格式
        details = parse_error_details(result.error or "")
        assert details is not None
        assert details["format"] == "image/tiff"
        assert "image/png" in details["accepted"]
    finally:
        await client.disconnect()
