"""
Contract Tests for word:select:text

测试 word:select:text 事件的完整请求-响应流程。

Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/42467331
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_success(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试成功选中文本的完整流程。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册响应
    3. 连接到 Workspace
    4. Workspace 发送 word:select:text 命令
    5. 验证返回的响应数据符合协议
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_match_count = 3
    expected_selected_index = 2
    expected_text = "Hello World"

    def select_text_response(request: dict) -> dict:
        """动态响应工厂，验证请求数据并返回响应"""
        # 验证请求数据
        assert "requestId" in request
        assert "documentUri" in request
        assert request["searchText"] == expected_text
        assert request["selectionMode"] == "select"
        assert request["selectIndex"] == expected_selected_index

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.select_text_response(
                success=True,
                match_count=expected_match_count,
                selected_index=expected_selected_index,
                selected_text=expected_text,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    # 创建客户端
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    # 在连接之前注册响应
    client.register_response("word:select:text", select_text_response)

    # 连接到服务器
    await client.connect()

    try:
        # Act: 执行动作
        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": client.document_uri,
                "searchText": expected_text,
                "selectionMode": "select",
                "selectIndex": expected_selected_index,
            },
        )
        result = await workspace.execute(action)

        # Assert: 验证结果
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["matchCount"] == expected_match_count
        assert result.data["selectedIndex"] == expected_selected_index
        assert result.data["selectedText"] == expected_text
        assert result.data["success"] is True

        # 验证 Mock 客户端接收到了请求
        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:select:text"
        assert event_data["searchText"] == expected_text
        assert event_data["selectionMode"] == "select"
        assert event_data["selectIndex"] == expected_selected_index
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_with_search_options(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试带搜索选项的文本选择。

    验证 matchCase、matchWholeWord、matchWildcards 选项正确传递。
    """
    # Arrange
    search_text = "test pattern"

    def select_text_response(request: dict) -> dict:
        # 验证搜索选项
        assert "searchOptions" in request
        assert request["searchOptions"]["matchCase"] is True
        assert request["searchOptions"]["matchWholeWord"] is True
        assert request["searchOptions"]["matchWildcards"] is False

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.select_text_response(
                success=True,
                match_count=1,
                selected_index=1,
                selected_text=search_text,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:select:text", select_text_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": client.document_uri,
                "searchText": search_text,
                "searchOptions": {
                    "matchCase": True,
                    "matchWholeWord": True,
                    "matchWildcards": False,
                },
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["matchCount"] == 1

        # 验证搜索选项被正确发送
        _, event_data = client.received_events[0]
        assert event_data["searchOptions"]["matchCase"] is True
        assert event_data["searchOptions"]["matchWholeWord"] is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_different_modes(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试不同的选择模式（select/start/end）。

    验证 selectionMode 参数正确传递。
    """
    # Arrange
    modes = ["select", "start", "end"]

    for mode in modes:

        def select_text_response(request: dict, m: list[str] = mode) -> dict:
            # 验证选择模式
            assert request["selectionMode"] == m

            return {
                "requestId": request["requestId"],
                "success": True,
                "data": word_factory.select_text_response(
                    success=True,
                    match_count=1,
                    selected_index=1,
                    selected_text="test",
                ),
                "timestamp": int(asyncio.get_event_loop().time() * 1000),
            }

        client = mock_word_client_factory(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id=f"contract_test_word_client_{mode}",
            document_uri="file:///tmp/contract_test.docx",
        )

        client.register_response("word:select:text", select_text_response)
        await client.connect()

        try:
            # Act
            action = OfficeAction(
                category="word",
                action_name="select:text",
                params={
                    "document_uri": client.document_uri,
                    "searchText": "test",
                    "selectionMode": mode,
                },
            )
            result = await workspace.execute(action)

            # Assert
            assert result.success is True

            # 验证选择模式被正确发送
            _, event_data = client.received_events[0]
            assert event_data["selectionMode"] == mode
        finally:
            await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_no_matches(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试未找到匹配文本的情况。

    验证 matchCount = 0 时返回的响应。

    ⚠️ **规范层反例守护，勿改成错误用例**（oasp#23 §SEARCH_NO_MATCH / issue #90）：
    ``matchCount: 0`` 是**良定义的零元结果**，MUST 返回 ``success: true``，**不得**返回
    ``3012 SEARCH_NO_MATCH``。该区分在线缆上可观测——本事件规范早已把 ``matchCount: 0``
    写成正常成功响应。对照 ``test_insert_comment.py::test_insert_comment_search_no_match_3012``。
    """
    # Arrange
    search_text = "nonexistent text"

    def select_text_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.select_text_response(
                success=True,
                match_count=0,
                selected_index=1,
                selected_text="",
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:select:text", select_text_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": client.document_uri,
                "searchText": search_text,
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["matchCount"] == 0
        assert result.data["selectedText"] == ""
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_select_nth_match(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试选择第N个匹配项。

    验证 selectIndex 参数正确工作，能够选择非第一个匹配项。
    """
    # Arrange
    total_matches = 5
    select_index = 3

    def select_text_response(request: dict) -> dict:
        # 验证 selectIndex
        assert request["selectIndex"] == select_index

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.select_text_response(
                success=True,
                match_count=total_matches,
                selected_index=select_index,
                selected_text=f"Match {select_index}",
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:select:text", select_text_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "test",
                "selectIndex": select_index,
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["matchCount"] == total_matches
        assert result.data["selectedIndex"] == select_index
        assert result.data["selectedText"] == f"Match {select_index}"

        # 验证 selectIndex 被正确发送
        _, event_data = client.received_events[0]
        assert event_data["selectIndex"] == select_index
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_with_wildcards(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试通配符搜索。

    验证 matchWildcards 选项正确传递和处理。
    """
    # Arrange
    search_pattern = "test*"

    def select_text_response(request: dict) -> dict:
        # 验证通配符选项
        assert "searchOptions" in request
        assert request["searchOptions"]["matchWildcards"] is True

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.select_text_response(
                success=True,
                match_count=4,
                selected_index=1,
                selected_text="test pattern",
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:select:text", select_text_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": client.document_uri,
                "searchText": search_pattern,
                "searchOptions": {
                    "matchWildcards": True,
                },
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["matchCount"] == 4

        # 验证通配符选项被正确发送
        _, event_data = client.received_events[0]
        assert event_data["searchText"] == search_pattern
        assert event_data["searchOptions"]["matchWildcards"] is True
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_select_text_error_response(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试错误响应。

    验证当操作失败时返回正确的错误信息。
    """
    # Arrange

    def select_text_response(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3000",
                "message": "Office API error",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:select:text", select_text_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="select:text",
            params={
                "document_uri": client.document_uri,
                "searchText": "test",
            },
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert result.error is not None
        # Note: error is a string in OfficeResult
        assert "Office API error" in str(result.error)
    finally:
        await client.disconnect()
