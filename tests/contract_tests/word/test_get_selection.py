"""
Contract Tests for word:get:selection

测试 word:get:selection 事件的完整请求-响应流程。
"""

from __future__ import annotations

import asyncio

import pytest

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selection_normal(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试获取正常选区（有高亮文本）的完整流程。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册正常选区响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:selection 命令
    5. 验证返回的选区数据符合协议
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_text = "Selected text from contract test"
    expected_start = 100
    expected_end = 130

    def selection_response(request: dict) -> dict:
        """动态响应工厂，验证请求数据并返回响应"""
        # 验证请求数据
        assert "requestId" in request
        assert "documentUri" in request

        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.selection_response(
                is_empty=False,
                selection_type="Normal",
                start=expected_start,
                end=expected_end,
                text=expected_text,
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
    client.register_response("word:get:selection", selection_response)

    # 连接到服务器
    await client.connect()

    try:
        # Act: 执行动作
        action = OfficeAction(
            category="word",
            action_name="get:selection",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert: 验证结果
        assert result.success is True, f"Expected success, got error: {result.error}"
        assert result.data["isEmpty"] is False
        assert result.data["type"] == "Normal"
        assert result.data["start"] == expected_start
        assert result.data["end"] == expected_end
        assert result.data["text"] == expected_text

        # 验证 Mock 客户端接收到了请求
        assert len(client.received_events) == 1
        event_name, event_data = client.received_events[0]
        assert event_name == "word:get:selection"
        assert "requestId" in event_data
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selection_insertion_point(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试获取光标位置（插入点，无高亮文本）。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册光标位置响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:selection 命令
    5. 验证光标位置数据正确返回
    """
    # Arrange: 创建并配置 Mock 客户端
    expected_position = 250

    def insertion_point_response(request: dict) -> dict:
        """光标位置响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.selection_response(
                is_empty=True,
                selection_type="InsertionPoint",
                start=expected_position,
                end=expected_position,
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:selection", insertion_point_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selection",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["isEmpty"] is True
        assert result.data["type"] == "InsertionPoint"
        assert result.data["start"] == expected_position
        assert result.data["end"] == expected_position
        assert result.data.get("text") is None
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selection_no_selection(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试无选区状态。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册无选区响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:selection 命令
    5. 验证无选区状态正确返回
    """
    # Arrange
    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_static_response(
        "word:get:selection",
        {
            "requestId": "test_req_001",
            "success": True,
            "data": word_factory.selection_response(
                is_empty=True,
                selection_type="NoSelection",
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        },
    )

    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selection",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is True
        assert result.data["isEmpty"] is True
        assert result.data["type"] == "NoSelection"
        assert result.data.get("start") is None
        assert result.data.get("end") is None
        assert result.data.get("text") is None
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selection_error(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
):
    """
    测试获取选区时的错误处理。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册错误响应
    3. 连接到 Workspace
    4. Workspace 发送 word:get:selection 命令
    5. 验证错误信息正确返回
    """

    # Arrange: 创建并配置 Mock 客户端
    def error_response(request: dict) -> dict:
        """错误响应工厂"""
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3000",
                "message": "DOCUMENT_ERROR - Failed to get selection",
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:selection", error_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selection",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        assert result.success is False
        assert "3000" in str(result.error) or "office" in str(result.error).lower()
    finally:
        await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_get_selection_timeout(
    workspace: OfficeWorkspace,
    mock_word_client_factory,
    word_factory,
):
    """
    测试请求超时场景。

    测试步骤：
    1. 创建 Mock Add-In 客户端
    2. 注册延迟响应
    3. 连接到 Workspace
    4. Workspace 发送命令
    5. 验证超时错误
    """

    # Arrange: 临时缩短超时时间，避免测试等待过久
    original_timeout = workspace.config.request_timeout
    workspace.config.request_timeout = 1000  # 1 秒

    # 注册延迟响应
    async def slow_response(request: dict) -> dict:
        """延迟响应工厂"""
        await asyncio.sleep(2)  # 超过测试超时时间（1秒）
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": word_factory.selection_response(
                is_empty=False,
                selection_type="Normal",
                start=0,
                end=10,
                text="Delayed",
            ),
        }

    client = mock_word_client_factory(
        server_url="http://127.0.0.1:3003",
        namespace="/word",
        client_id="contract_test_word_client",
        document_uri="file:///tmp/contract_test.docx",
    )

    client.register_response("word:get:selection", slow_response)
    await client.connect()

    try:
        # Act
        action = OfficeAction(
            category="word",
            action_name="get:selection",
            params={"document_uri": client.document_uri},
        )
        result = await workspace.execute(action)

        # Assert
        if result.success:
            pytest.fail(f"Expected timeout error, but got success: {result.data}")
        # 验证有错误信息
        assert result.error is not None or result.metadata.get("error"), (
            f"Expected error, got: error='{result.error}', metadata={result.metadata}"
        )
    finally:
        workspace.config.request_timeout = original_timeout
        await client.disconnect()
