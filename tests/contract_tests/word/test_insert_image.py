"""
Contract Tests for word:insert:image

测试 word:insert:image 事件的完整请求-响应流程。

本文件随 issue #90（oasp#23 四孤儿码落地）新建——该事件此前**无任何契约测试**，
故 3007 输入侧判法无处落地。补齐成功路径 + 4002/3007 二分。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace, parse_error_details

# 1x1 透明 PNG，最小可解码载荷
TEST_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="


def _client(mock_word_client_factory):
    return mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )


def _err(request: dict, code: str, message: str, details: dict | None = None) -> dict:
    error: dict = {"code": code, "message": message}
    if details:
        error["details"] = details
    return {
        "requestId": request["requestId"],
        "success": False,
        "error": error,
        "timestamp": int(asyncio.get_event_loop().time() * 1000),
    }


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_image_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """成功插入图片，返回 inserted + imageId。"""

    def response_factory(request: dict) -> dict:
        assert request["image"]["base64"] == TEST_BASE64

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": {"inserted": True, "imageId": "shape-12345"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = _client(mock_word_client_factory)
    client.register_response("word:insert:image", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="insert:image",
            params={
                "document_uri": client.document_uri,
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["inserted"] is True
        assert result.data["imageId"] == "shape-12345"

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:insert:image"
        assert event_data["image"]["base64"] == TEST_BASE64
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_image_undecodable_base64_4002(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """base64 在线缆层就解不开 → 4002 INVALID_PARAM（**不是** 3007）。

    与 3007 成对，钉住二分的**另一半**：只有两条都在，才证明划界被真正实现而非碰巧命中。
    """

    def error_response(request: dict) -> dict:
        return _err(request, "4002", "image.base64 is not valid base64")

    client = _client(mock_word_client_factory)
    client.register_response("word:insert:image", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="insert:image",
            params={
                "document_uri": client.document_uri,
                "image": {"base64": "!!!not-base64!!!"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "4002" in (result.error or "")
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_image_format_not_supported_3007(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """base64 可解码但图片格式不受支持 → 3007 FORMAT_NOT_SUPPORTED（oasp#23 / issue #90）。

    输入侧语义：数据完整、解得开，只是宿主拒绝嵌入这一种格式——恢复动作是**转码后重发**
    （如转 PNG），而非修参数。原落 3000 兜底时调用方无从区分「文档操作失败」与「换个格式就行」。
    """

    def error_response(request: dict) -> dict:
        return _err(
            request,
            "3007",
            "Image format 'image/tiff' is not supported",
            {"format": "image/tiff", "accepted": ["image/png", "image/jpeg"]},
        )

    client = _client(mock_word_client_factory)
    client.register_response("word:insert:image", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="insert:image",
            params={
                "document_uri": client.document_uri,
                "image": {"base64": TEST_BASE64},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3007" in (result.error or "")
        # details 回带被拒格式与可接受列表，调用方据此一次选对转码目标
        details = parse_error_details(result.error or "")
        assert details is not None
        assert details["format"] == "image/tiff"
        assert "image/png" in details["accepted"]
    finally:
        await client.disconnect()
