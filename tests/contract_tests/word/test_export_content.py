"""
Contract Tests for word:export:content

测试 word:export:content 事件的完整请求-响应流程。

本文件随 issue #90（oasp#23 四孤儿码落地）新建——该事件此前**无任何契约测试**，
故 3007 输出侧判法与其降级路由 3016 无处落地。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace, parse_error_details


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
async def test_export_content_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """成功导出 markdown，返回 content + format。"""

    def response_factory(request: dict) -> dict:
        assert request["format"] == "markdown"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": {"content": "# 标题\n\n这是文档内容...", "format": "markdown"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = _client(mock_word_client_factory)
    client.register_response("word:export:content", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="export:content",
            params={"document_uri": client.document_uri, "format": "markdown"},
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["format"] == "markdown"
        assert "标题" in result.data["content"]

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:export:content"
        assert event_data["format"] == "markdown"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_export_content_format_not_supported_3007(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """format 枚举合法但目标产不出该格式 → 3007 FORMAT_NOT_SUPPORTED（oasp#23 / issue #90）。

    输出侧语义：``markdown`` 是 DTO 枚举内的合法取值（故过得了线缆层校验），只是这个宿主
    导不出——恢复动作是**改请求另一格式重发**（如退到 ``html``）。原落 3000 兜底。
    """

    def error_response(request: dict) -> dict:
        return _err(
            request,
            "3007",
            "Export format 'markdown' is not supported by this host",
            {"format": "markdown", "accepted": ["text", "html"]},
        )

    client = _client(mock_word_client_factory)
    client.register_response("word:export:content", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="export:content",
            params={"document_uri": client.document_uri, "format": "markdown"},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3007" in (result.error or "")
        # accepted 列表让调用方一次选对退路，不必逐格式试错
        details = parse_error_details(result.error or "")
        assert details is not None
        assert details["format"] == "markdown"
        assert "html" in details["accepted"]
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_export_content_api_not_supported_3016(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """整个导出能力在当前宿主不可用 → 3016 API_NOT_SUPPORTED（oasp#23 / issue #90）。

    与 3007 的分界是**失败粒度**：3007 说「换个格式就行」，3016 说「换格式无用，须换路径或
    平台」。导出高度依赖宿主能力，故 3016 是其降级路由、在规范里占**独立一行**。
    """

    def error_response(request: dict) -> dict:
        return _err(request, "3016", "Content export is not available on this platform")

    client = _client(mock_word_client_factory)
    client.register_response("word:export:content", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="export:content",
            params={"document_uri": client.document_uri, "format": "text"},
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3016" in (result.error or "")
    finally:
        await client.disconnect()
