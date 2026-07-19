"""
Contract Tests for word:insert:comment

测试 word:insert:comment 事件的完整请求-响应流程。

本文件随 issue #90（oasp#23 四孤儿码落地）新建——该事件此前**无任何契约测试**，
而它是 `3012 SEARCH_NO_MATCH` 规范层判法的**唯一** MUST 归属事件。
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


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_comment_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """searchText 命中 → 成功附加批注，返回 commentId。"""

    def response_factory(request: dict) -> dict:
        assert request["text"] == "需要复核"
        assert request["target"]["type"] == "searchText"
        assert request["target"]["searchText"] == "季度营收"

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": {"commentId": "comment-001"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = _client(mock_word_client_factory)
    client.register_response("word:insert:comment", response_factory)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="insert:comment",
            params={
                "document_uri": client.document_uri,
                "text": "需要复核",
                "target": {"type": "searchText", "searchText": "季度营收"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["commentId"] == "comment-001"

        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:insert:comment"
        assert event_data["target"]["searchText"] == "季度营收"
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_comment_search_no_match_3012(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """searchText 零匹配 → 3012 SEARCH_NO_MATCH（**MUST**，oasp#23 规范层 / issue #90）。

    判据是「零匹配是否仍是一个**良定义的零元结果**」：把批注附到 0 个范围上**没有**对应的
    良定义结果——请求什么也没做成，故属前置条件不满足，MUST 报 3012。

    原实现降级为通用 ``3000 DOCUMENT_ERROR``，正是 error-handling.md §SEARCH_NO_MATCH (3012)
    规范层明令禁止的模式：
    调用方无法区分「文档操作失败」（通常应上报）与「搜索无匹配」（应放宽搜索选项重试，或改用
    ``type: "selection"`` 让用户先选中）——两者的恢复动作完全不同。

    反向守护（零元结果**不得**判 3012）见 ``test_replace_text.py`` / ``test_select_text.py`` /
    ``excel/test_find_filter.py``。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3012",
                "message": "No match found for the search text",
                "details": {"searchText": "不存在的文本", "matchCase": False, "matchWholeWord": False},
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = _client(mock_word_client_factory)
    client.register_response("word:insert:comment", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="insert:comment",
            params={
                "document_uri": client.document_uri,
                "text": "需要复核",
                "target": {"type": "searchText", "searchText": "不存在的文本"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3012" in (result.error or "")
        # details 回带 searchText 与实际生效的搜索选项，供调用方放宽条件后重试
        details = parse_error_details(result.error or "")
        assert details is not None
        assert details["searchText"] == "不存在的文本"
        assert details["matchCase"] is False
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_insert_comment_selection_empty_3002(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """selection 模式下无活动选区 → 3002 SELECTION_EMPTY（**不是** 3012）。

    与 3012 成对：3012 专属 ``searchText`` 模式，``selection`` 模式的「没东西可附」由 3002
    承载。两码混用会让调用方走错恢复路径（放宽搜索词 vs 提示用户先选中）。
    """

    def error_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {"code": "3002", "message": "No active selection"},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = _client(mock_word_client_factory)
    client.register_response("word:insert:comment", error_response)
    await client.connect()

    try:
        action = OfficeAction(
            category="word",
            action_name="insert:comment",
            params={
                "document_uri": client.document_uri,
                "text": "需要复核",
                "target": {"type": "selection"},
            },
        )
        result = await workspace.execute(action)

        assert result.success is False
        assert "3002" in (result.error or "")
    finally:
        await client.disconnect()
