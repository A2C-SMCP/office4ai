"""
Contract Tests for ppt:update:textBox

测试 ppt:update:textBox 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_text_box_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试成功更新文本框文本。"""

    def response_factory(request: dict) -> dict:
        assert "requestId" in request
        assert request["elementId"] == "shape-001"
        assert request["updates"]["text"] == "New content"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_text_box_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:textBox", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:textBox",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-001",
                "updates": {"text": "New content"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["updatedCount"] == 1

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "ppt:update:textBox"
        assert event_data["elementId"] == "shape-001"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_text_box_text_and_style(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    ppt_factory,
):
    """测试同时更新文本和样式。"""

    def response_factory(request: dict) -> dict:
        updates = request["updates"]
        assert updates["text"] == "Bold Title"
        # OASP 0.4.0: 字体收敛到 font 子对象
        assert updates["font"]["size"] == 32
        assert updates["font"]["bold"] is True
        assert updates["font"]["color"] == "#FF0000"
        assert updates["font"]["underline"] == "WavyHeavy"
        assert updates["fillColor"] == "#EEEEEE"
        # run 级局部格式 + 段落 bullet（嵌套区间寻址）
        assert updates["runs"][0] == {"start": 0, "length": 4, "font": {"color": "#C00000", "allCaps": True}}
        assert updates["paragraphs"][0]["bulletFormat"] == {
            "visible": True,
            "type": "Numbered",
            "style": "ArabicNumeralPeriod",
        }

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": ppt_factory.update_text_box_response(),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:textBox", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:textBox",
            params={
                "document_uri": client.document_uri,
                "elementId": "shape-001",
                "updates": {
                    "text": "Bold Title",
                    "font": {
                        "size": 32,
                        "name": "Arial",
                        "color": "#FF0000",
                        "bold": True,
                        "italic": False,
                        "underline": "WavyHeavy",
                    },
                    "fillColor": "#EEEEEE",
                    "runs": [{"start": 0, "length": 4, "font": {"color": "#C00000", "allCaps": True}}],
                    "paragraphs": [
                        {
                            "start": 0,
                            "length": 20,
                            "bulletFormat": {"visible": True, "type": "Numbered", "style": "ArabicNumeralPeriod"},
                        }
                    ],
                },
            },
        )
        result = await workspace.execute(action)

        assert result.success is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_update_text_box_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """测试错误处理。"""

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3001", "message": "Element not found"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/ppt",
        client_id="contract_test_ppt_client",
        document_uri="file:///tmp/test.pptx",
    )

    client.register_response("ppt:update:textBox", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="ppt",
            action_name="update:textBox",
            params={
                "document_uri": client.document_uri,
                "elementId": "nonexistent",
                "updates": {"text": "Test"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
    finally:
        await client.disconnect()
